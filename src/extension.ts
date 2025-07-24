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
            (uri: vscode.Uri) => handleBatchConvert(uri, 'all', context))
    ];
    
    context.subscriptions.push(...disposables);
}

export function deactivate() {
    console.log('Markdown Hub is now deactivated.');
}