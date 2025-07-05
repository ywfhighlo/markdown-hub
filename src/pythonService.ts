import * as vscode from 'vscode';
import * as path from 'path';
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

/**
 * 执行 Python 后端脚本进行文件转换
 */
export function executePythonScript(
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
        
        console.log(`执行命令: ${pythonPath} ${args.join(' ')}`);

        const pyProcess = spawn(pythonPath, args);

        // **关键修复**: 明确设置stdout和stderr的编码为utf8
        // 这可以防止在Windows上因默认编码不同而导致的中文乱码问题
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

 