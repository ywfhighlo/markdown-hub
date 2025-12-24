import * as vscode from 'vscode';
import { handleConvertCommand, handleOpenTemplateSettingsCommand } from './commandHandler';

// 批量转换处理函数 - 复用现有逻辑
async function handleBatchConvert(uri: vscode.Uri, fileType: string, context: vscode.ExtensionContext) {
    if (!uri) {
        vscode.window.showErrorMessage('请选择一个目录进行批量转换。');
        return;
    }

    // 设置文件类型过滤环境变量
    const originalEnv = process.env.BATCH_FILTER_TYPE;
    process.env.BATCH_FILTER_TYPE = fileType;
    
    try {
        // 复用现有的 office-to-md 转换逻辑
        await handleConvertCommand(uri, 'office-to-md', context);
    } finally {
        // 恢复环境变量
        if (originalEnv !== undefined) {
            process.env.BATCH_FILTER_TYPE = originalEnv;
        } else {
            delete process.env.BATCH_FILTER_TYPE;
        }
    }
}

import * as path from 'path';
import * as child_process from 'child_process';

// 批量PDF转PNG处理函数
async function handleBatchPdfToPng(uri: vscode.Uri, context: vscode.ExtensionContext) {
    if (!uri) {
        vscode.window.showErrorMessage('请选择一个目录进行批量转换。');
        return;
    }

    const scriptPath = path.join(context.extensionPath, 'backend', 'converters', 'batch_pdf_to_png.py');
    const targetDir = uri.fsPath;

    // 获取配置中的 python 路径
    const config = vscode.workspace.getConfiguration('markdownHub');
    const pythonPath = config.get<string>('pythonPath') || 'python';
    
    // 获取配置中的 poppler 路径
    let popplerPath = config.get<string>('popplerPath') || '';
    
    // 如果配置未设置，尝试使用内置的 poppler
    if (!popplerPath) {
        const localPopplerPath = path.join(context.extensionPath, 'tools', 'poppler', 'poppler-24.02.0', 'Library', 'bin');
        // 由于无法简单检测文件夹是否存在(fs需要import)，这里直接传递
        // 脚本端会校验
        popplerPath = localPopplerPath;
    }

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: "正在将单页PDF转换为PNG...",
        cancellable: false
    }, async (progress, token) => {
        return new Promise<void>((resolve, reject) => {
            const args = [scriptPath, targetDir];
            if (popplerPath) {
                args.push('--poppler-path', popplerPath);
            }

            const process = child_process.spawn(pythonPath, args);
            
            let output = '';
            let errorOutput = '';

            process.stdout.on('data', (data) => {
                const msg = data.toString();
                output += msg;
                // 简单的进度反馈
                if (msg.includes('Converting')) {
                    progress.report({ message: msg.trim() });
                }
            });

            process.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            process.on('close', (code) => {
                if (code === 0) {
                    vscode.window.showInformationMessage(`批量转换完成。请查看输出窗口了解详情。`);
                    resolve();
                } else {
                    vscode.window.showErrorMessage(`转换失败 (代码 ${code}): ${errorOutput || output}`);
                    resolve(); // Resolve anyway to close progress
                }
            });
        });
    });
}

export function activate(context: vscode.ExtensionContext) {
    console.log('Markdown Hub is now active!');
    
    // Register all conversion commands
    const disposables = [
        vscode.commands.registerCommand('markdown-hub.mdToDocx', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-docx', context)),
        
        vscode.commands.registerCommand('markdown-hub.mdToPdf', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-pdf', context)),
        
        vscode.commands.registerCommand('markdown-hub.mdToHtml', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-html', context)),
        
        vscode.commands.registerCommand('markdown-hub.mdToPptx', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-pptx', context)),
        
        vscode.commands.registerCommand('markdown-hub.officeToMd', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'office-to-md', context)),
        
        vscode.commands.registerCommand('markdown-hub.diagramToPng', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'diagram-to-png', context)),
        
        vscode.commands.registerCommand('markdown-hub.openTemplateSettings', 
            () => handleOpenTemplateSettingsCommand()),
        
        // 注册新的批量转换命令
        vscode.commands.registerCommand('markdown-hub.batchMdToPdf', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-pdf', context)),
            
        vscode.commands.registerCommand('markdown-hub.batchMdToDocx', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-docx', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchMdToPptx', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'md-to-pptx', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchPdfToMd', 
            (uri: vscode.Uri) => handleBatchConvert(uri, 'pdf', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchDocxToMd', 
            (uri: vscode.Uri) => handleBatchConvert(uri, 'docx', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchPptxToMd', 
            (uri: vscode.Uri) => handleBatchConvert(uri, 'pptx', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchExcelToMd', 
            (uri: vscode.Uri) => handleBatchConvert(uri, 'excel', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchAllToMd', 
            (uri: vscode.Uri) => handleBatchConvert(uri, 'all', context)),
        
        vscode.commands.registerCommand('markdown-hub.batchDiagramToPng', 
            (uri: vscode.Uri) => handleConvertCommand(uri, 'diagram-to-png', context)),
            
        vscode.commands.registerCommand('markdown-hub.batchPdfToPng', 
            (uri: vscode.Uri) => handleBatchPdfToPng(uri, context))
    ];
    
    context.subscriptions.push(...disposables);
}

export function deactivate() {
    console.log('Markdown Hub is now deactivated.');
}