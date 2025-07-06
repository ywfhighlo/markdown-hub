import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { spawn } from 'child_process';

type ConversionType = 'md-to-docx' | 'md-to-pdf' | 'md-to-html' | 'md-to-pptx' | 'office-to-md' | 'diagram-to-png';

interface PythonResponse {
    success: boolean;
    outputFiles?: string[];
    error?: string;
}

interface ProgressInfo {
    type: 'progress';
    stage: string;
    percentage?: number;
}

interface ResultInfo {
    type: 'result';
    success: boolean;
    outputFiles?: string[];
    error?: string;
}

type PythonOutput = ProgressInfo | ResultInfo;

interface BundledDepsInfo {
    bundledAt: string;
    hasFullDeps: boolean;
    hasCairo: boolean;
    pythonDepsPath: string;
    backendPath: string;
    note: string;
}

/**
 * 检查是否有打包的依赖
 */
function checkBundledDeps(context: vscode.ExtensionContext): BundledDepsInfo | null {
    try {
        const depsInfoPath = path.join(context.extensionPath, 'bundled', 'deps-info.json');
        if (fs.existsSync(depsInfoPath)) {
            const depsInfo = JSON.parse(fs.readFileSync(depsInfoPath, 'utf8')) as BundledDepsInfo;
            
            // 验证打包的依赖是否存在
            const pythonDepsPath = path.join(context.extensionPath, depsInfo.pythonDepsPath);
            const backendPath = path.join(context.extensionPath, depsInfo.backendPath);
            
            if (fs.existsSync(pythonDepsPath) && fs.existsSync(backendPath)) {
                console.log('✓ 检测到打包的 Python 依赖');
                console.log(`  - 功能状态: ${depsInfo.hasFullDeps ? '完整功能' : '基础功能'}`);
                console.log(`  - Cairo 支持: ${depsInfo.hasCairo ? '是' : '否'}`);
                return depsInfo;
            }
        }
    } catch (error) {
        console.log('检查打包依赖时出错:', error);
    }
    
    return null;
}

/**
 * 使用打包的依赖执行 Python 脚本
 */
function executePythonWithBundledDeps(
    sourcePath: string,
    conversionType: ConversionType,
    outputDir: string,
    context: vscode.ExtensionContext,
    bundledDepsInfo: BundledDepsInfo,
    conversionOptions?: any,
    progressCallback?: (message: string, percentage?: number) => void
): Promise<PythonResponse> {
    
    return new Promise((resolve, reject) => {
        // 使用打包的后端脚本
        const scriptPath = path.join(context.extensionPath, bundledDepsInfo.backendPath, 'cli.py');
        const pythonDepsPath = path.join(context.extensionPath, bundledDepsInfo.pythonDepsPath);
        
        // 获取 Python 路径配置
        const config = vscode.workspace.getConfiguration('markdown-hub');
        const isWindows = process.platform === 'win32';
        const defaultPythonCommand = isWindows ? 'python' : 'python3';
        const pythonPath = config.get<string>('pythonPath', defaultPythonCommand);

        const args = [
            scriptPath,
            '--conversion-type', conversionType,
            '--input-path', sourcePath,
            '--output-dir', outputDir
        ];
        
        // 添加转换选项参数
        if (conversionOptions) {
            // DOCX模板参数
            if (conversionOptions.docxTemplatePath) {
                args.push('--docx-template-path', conversionOptions.docxTemplatePath);
            }

            // PPTX模板参数
            if (conversionOptions.pptxTemplatePath) {
                args.push('--pptx-template-path', conversionOptions.pptxTemplatePath);
            }

            // 添加项目信息参数
            if (conversionOptions.projectName) {
                args.push('--project-name', conversionOptions.projectName);
            }
            if (conversionOptions.author) {
                args.push('--author', conversionOptions.author);
            }
            if (conversionOptions.email) {
                args.push('--email', conversionOptions.email);
            }
            if (conversionOptions.mobilephone) {
                args.push('--mobilephone', conversionOptions.mobilephone);
            }
            
            // 添加标题提升参数
            if (conversionOptions.promoteHeadings) {
                args.push('--promote-headings');
            }

            // 添加 Poppler 路径参数
            if (conversionOptions.popplerPath) {
                args.push('--poppler-path', conversionOptions.popplerPath);
            }

            // 添加 Tesseract 命令/路径参数
            if (conversionOptions.tesseractCmd) {
                args.push('--tesseract-cmd', conversionOptions.tesseractCmd);
            }
        }
        
        console.log(`使用打包依赖执行命令: ${pythonPath} ${args.join(' ')}`);

        // 设置环境变量，使用打包的依赖
        const env = { ...process.env };
        const pathSeparator = process.platform === 'win32' ? ';' : ':';
        env.PYTHONPATH = pythonDepsPath + pathSeparator + (env.PYTHONPATH || '');

        const pyProcess = spawn(pythonPath, args, { env });

        // **关键修复**: 明确设置stdout和stderr的编码为utf8
        pyProcess.stdout.setEncoding('utf8');
        pyProcess.stderr.setEncoding('utf8');

        let stdoutBuffer = '';

        pyProcess.stdout.on('data', (data) => {
            stdoutBuffer += data;
            
            // 尝试处理缓冲区中的每一行
            let lines = stdoutBuffer.split('\n');
            // 保留最后一个不完整的行（如果有）
            stdoutBuffer = lines[lines.length - 1];
            // 处理完整的行
            lines.slice(0, -1).forEach(line => {
                try {
                    // Base64解码
                    const decodedLine = Buffer.from(line, 'base64').toString('utf8');
                    const output = JSON.parse(decodedLine) as PythonOutput;

                    if (output.type === 'progress' && progressCallback) {
                        progressCallback(output.stage, output.percentage);
                    } else if (output.type === 'result') {
                        if (output.success) {
                            resolve({
                                success: true,
                                outputFiles: output.outputFiles
                            });
                        } else {
                            reject(new Error(output.error || 'Python 脚本报告了一个未知错误'));
                        }
                    }
                } catch (e) {
                    console.log('非JSON输出:', line);
                }
            });
        });

        pyProcess.stderr.on('data', (data) => {
            console.error(`Python错误输出: ${data}`);
        });

        pyProcess.on('close', (code) => {
            if (code !== 0 && stdoutBuffer.trim()) {
                // 如果进程异常退出且还有未处理的输出，尝试解析
                try {
                    const decodedLine = Buffer.from(stdoutBuffer.trim(), 'base64').toString('utf8');
                    const finalOutput = JSON.parse(decodedLine) as PythonOutput;
                    if (finalOutput.type === 'result' && !finalOutput.success) {
                        const errorMessage = finalOutput.error || `Python 脚本异常退出，代码：${code}`;
                        reject(new Error(errorMessage));
                        return;
                    }
                } catch (e) {
                    // 解析失败，使用通用错误信息
                    reject(new Error(`Python 脚本异常退出，代码：${code}。输出：${stdoutBuffer.trim()}`));
                }
            } else if (code !== 0) {
                 reject(new Error(`Python 脚本异常退出，代码：${code}`));
            }
        });

        pyProcess.on('error', (err) => {
            reject(new Error(`无法启动 Python 脚本：${err.message}`));
        });
    });
}

/**
 * 使用用户环境执行 Python 脚本（原有逻辑）
 */
function executePythonWithUserEnv(
    sourcePath: string,
    conversionType: ConversionType,
    outputDir: string,
    context: vscode.ExtensionContext,
    conversionOptions?: any,
    progressCallback?: (message: string, percentage?: number) => void
): Promise<PythonResponse> {
    
    return new Promise((resolve, reject) => {
        // Python 脚本路径
        const scriptPath = path.join(context.extensionPath, 'backend', 'cli.py');
        
        // 获取 Python 路径配置，并根据操作系统智能选择默认值
        const config = vscode.workspace.getConfiguration('markdown-hub');
        const isWindows = process.platform === 'win32';
        const defaultPythonCommand = isWindows ? 'python' : 'python3';
        const pythonPath = config.get<string>('pythonPath', defaultPythonCommand);

        const args = [
            scriptPath,
            '--conversion-type', conversionType,
            '--input-path', sourcePath,
            '--output-dir', outputDir
        ];
        
        // 添加转换选项参数
        if (conversionOptions) {
            // DOCX模板参数
            if (conversionOptions.docxTemplatePath) {
                args.push('--docx-template-path', conversionOptions.docxTemplatePath);
            }

            // PPTX模板参数
            if (conversionOptions.pptxTemplatePath) {
                args.push('--pptx-template-path', conversionOptions.pptxTemplatePath);
            }

            // 添加项目信息参数
            if (conversionOptions.projectName) {
                args.push('--project-name', conversionOptions.projectName);
            }
            if (conversionOptions.author) {
                args.push('--author', conversionOptions.author);
            }
            if (conversionOptions.email) {
                args.push('--email', conversionOptions.email);
            }
            if (conversionOptions.mobilephone) {
                args.push('--mobilephone', conversionOptions.mobilephone);
            }
            
            // 添加标题提升参数
            if (conversionOptions.promoteHeadings) {
                args.push('--promote-headings');
            }

            // 添加 Poppler 路径参数
            if (conversionOptions.popplerPath) {
                args.push('--poppler-path', conversionOptions.popplerPath);
            }

            // 添加 Tesseract 命令/路径参数
            if (conversionOptions.tesseractCmd) {
                args.push('--tesseract-cmd', conversionOptions.tesseractCmd);
            }
        }
        
        console.log(`使用用户环境执行命令: ${pythonPath} ${args.join(' ')}`);

        const pyProcess = spawn(pythonPath, args);

        // **关键修复**: 明确设置stdout和stderr的编码为utf8
        pyProcess.stdout.setEncoding('utf8');
        pyProcess.stderr.setEncoding('utf8');

        let stdoutBuffer = '';

        pyProcess.stdout.on('data', (data) => {
            stdoutBuffer += data;
            
            // 尝试处理缓冲区中的每一行
            let lines = stdoutBuffer.split('\n');
            // 保留最后一个不完整的行（如果有）
            stdoutBuffer = lines[lines.length - 1];
            // 处理完整的行
            lines.slice(0, -1).forEach(line => {
                try {
                    // Base64解码
                    const decodedLine = Buffer.from(line, 'base64').toString('utf8');
                    const output = JSON.parse(decodedLine) as PythonOutput;

                    if (output.type === 'progress' && progressCallback) {
                        progressCallback(output.stage, output.percentage);
                    } else if (output.type === 'result') {
                        if (output.success) {
                            resolve({
                                success: true,
                                outputFiles: output.outputFiles
                            });
                        } else {
                            reject(new Error(output.error || 'Python 脚本报告了一个未知错误'));
                        }
                    }
                } catch (e) {
                    console.log('非JSON输出:', line);
                }
            });
        });

        pyProcess.stderr.on('data', (data) => {
            console.error(`Python错误输出: ${data}`);
        });

        pyProcess.on('close', (code) => {
            if (code !== 0 && stdoutBuffer.trim()) {
                // 如果进程异常退出且还有未处理的输出，尝试解析
                try {
                    const decodedLine = Buffer.from(stdoutBuffer.trim(), 'base64').toString('utf8');
                    const finalOutput = JSON.parse(decodedLine) as PythonOutput;
                    if (finalOutput.type === 'result' && !finalOutput.success) {
                        const errorMessage = finalOutput.error || `Python 脚本异常退出，代码：${code}`;
                        reject(new Error(errorMessage));
                        return;
                    }
                } catch (e) {
                    // 解析失败，使用通用错误信息
                    reject(new Error(`Python 脚本异常退出，代码：${code}。输出：${stdoutBuffer.trim()}`));
                }
            } else if (code !== 0) {
                 reject(new Error(`Python 脚本异常退出，代码：${code}`));
            }
        });

        pyProcess.on('error', (err) => {
            reject(new Error(`无法启动 Python 脚本：${err.message}`));
        });
    });
}

/**
 * 执行 Python 后端脚本进行文件转换
 * 优先使用打包的依赖，如果没有则使用用户环境
 */
export function executePythonScript(
    sourcePath: string,
    conversionType: ConversionType,
    outputDir: string,
    context: vscode.ExtensionContext,
    conversionOptions?: any,
    progressCallback?: (message: string, percentage?: number) => void
): Promise<PythonResponse> {
    
    // 检查是否有打包的依赖
    const bundledDepsInfo = checkBundledDeps(context);
    
    if (bundledDepsInfo) {
        // 使用打包的依赖
        return executePythonWithBundledDeps(
            sourcePath, 
            conversionType, 
            outputDir, 
            context, 
            bundledDepsInfo, 
            conversionOptions, 
            progressCallback
        );
    } else {
        // 回退到用户环境
        console.log('未检测到打包的依赖，使用用户环境');
        return executePythonWithUserEnv(
            sourcePath, 
            conversionType, 
            outputDir, 
            context, 
            conversionOptions, 
            progressCallback
        );
    }
}

 