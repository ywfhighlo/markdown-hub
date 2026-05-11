import * as vscode from 'vscode';
import * as path from 'path';
import { spawn } from 'child_process';
import { execSync } from 'child_process';

type ConversionType = 'md-to-docx' | 'md-to-pdf' | 'md-to-html' | 'md-to-pptx' | 'md-svg-to-docx' | 'office-to-md' | 'diagram-to-png';

interface PythonResponse {
    success: boolean;
    outputFiles?: string[];
    error?: string;
}

interface ProgressInfo {
    type: 'progress';
    stage: string;
    percentage?: number;
    details?: {
        currentFile?: number;
        totalFiles?: number;
        currentFileName?: string;
        fileSize?: number;
        pageCount?: number;
    };
}

interface ResultInfo {
    type: 'result';
    success: boolean;
    outputFiles?: string[];
    error?: string;
}

interface ConversionStats {
    totalFiles?: number;
    currentFile?: number;
    fileSize?: number;
    pageCount?: number;
}

type PythonOutput = ProgressInfo | ResultInfo;

const STAGE_LABELS: Record<string, string> = {
    'parsing': '正在解析文件...',
    'extracting': '正在提取文本...',
    'processing_images': '正在处理图片...',
    'converting': '正在转换格式...',
    'rendering': '正在渲染内容...',
    'generating': '正在生成文件...',
    'saving': '正在保存文件...',
    'complete': '转换完成',
    'error': '处理出错',
    'preparing': '正在准备...',
    'analyzing': '正在分析内容...',
    'optimizing': '正在优化输出...',
    'exporting': '正在导出...'
};

function getStageLabel(stage: string): string {
    return STAGE_LABELS[stage] || stage;
}

function detectPythonCommand(): string {
    const candidates = ['python3', 'python'];

    for (const candidate of candidates) {
        try {
            execSync(`${candidate} --version`, { stdio: 'ignore' });
            return candidate;
        } catch (error) {
            continue;
        }
    }

    const isWindows = process.platform === 'win32';
    return isWindows ? 'python' : 'python3';
}

export function executePythonScript(
    sourcePath: string,
    conversionType: ConversionType,
    outputDir: string,
    context: vscode.ExtensionContext,
    conversionOptions?: any,
    progressCallback?: (message: string, percentage?: number, stats?: ConversionStats) => void
): Promise<PythonResponse> {

    return new Promise((resolve, reject) => {
        const scriptPath = path.join(context.extensionPath, 'backend', 'cli.py');

        const config = vscode.workspace.getConfiguration('markdown-hub');
        const configuredPythonPath = config.get<string>('pythonPath');

        const pythonPath = configuredPythonPath || detectPythonCommand();

        const args = [
            scriptPath,
            '--conversion-type', conversionType,
            '--input-path', sourcePath,
            '--output-dir', outputDir
        ];

        if (conversionOptions) {
            if (conversionOptions.docxTemplatePath) {
                args.push('--docx-template-path', conversionOptions.docxTemplatePath);
            }

            if (conversionOptions.pptxTemplatePath) {
                args.push('--pptx-template-path', conversionOptions.pptxTemplatePath);
            }

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

            if (conversionOptions.promoteHeadings) {
                args.push('--promote-headings');
            }

            if (conversionOptions.popplerPath) {
                args.push('--poppler-path', conversionOptions.popplerPath);
            }

            if (conversionOptions.tesseractCmd) {
                args.push('--tesseract-cmd', conversionOptions.tesseractCmd);
            }

            if (conversionOptions.svgDpi !== undefined) {
                args.push('--svg-dpi', conversionOptions.svgDpi.toString());
            }
            if (conversionOptions.svgOutputWidth !== undefined) {
                args.push('--svg-output-width', conversionOptions.svgOutputWidth.toString());
            }
        }

        console.log(`执行命令: ${pythonPath} ${args.join(' ')}`);

        const pyProcess = spawn(pythonPath, args);

        pyProcess.stdout.setEncoding('utf8');
        pyProcess.stderr.setEncoding('utf8');

        let stdoutBuffer = '';
        let lastProgressTime = 0;

        pyProcess.stdout.on('data', (data) => {
            stdoutBuffer += data;

            let lines = stdoutBuffer.split('\n');
            stdoutBuffer = lines[lines.length - 1];

            lines.slice(0, -1).forEach(line => {
                try {
                    const decodedLine = Buffer.from(line, 'base64').toString('utf8');
                    const output = JSON.parse(decodedLine) as PythonOutput;

                    if (output.type === 'progress' && progressCallback) {
                        const now = Date.now();
                        const timeSinceLastProgress = now - lastProgressTime;

                        if (timeSinceLastProgress >= 100 || output.percentage === 100) {
                            const stageLabel = getStageLabel(output.stage);
                            let message = stageLabel;

                            if (output.details?.currentFileName) {
                                message = `${stageLabel} (${path.basename(output.details.currentFileName)})`;
                            }

                            const stats: ConversionStats | undefined = output.details?.totalFiles ? {
                                totalFiles: output.details.totalFiles,
                                currentFile: output.details.currentFile,
                                fileSize: output.details.fileSize,
                                pageCount: output.details.pageCount
                            } : undefined;

                            progressCallback(message, output.percentage, stats);
                            lastProgressTime = now;
                        }
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
                try {
                    const decodedLine = Buffer.from(stdoutBuffer.trim(), 'base64').toString('utf8');
                    const finalOutput = JSON.parse(decodedLine) as PythonOutput;
                    if (finalOutput.type === 'result' && !finalOutput.success) {
                        const errorMessage = finalOutput.error || `Python 脚本异常退出，代码：${code}`;
                        reject(new Error(errorMessage));
                        return;
                    }
                } catch (e) {
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
