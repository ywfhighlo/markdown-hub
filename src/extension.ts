import * as vscode from 'vscode';
import { handleConvertCommand, handleOpenTemplateSettingsCommand, handleViewHistoryCommand, handleClearHistoryCommand, handleCheckDependenciesCommand } from './commandHandler';

// Batch conversion handler - reuses existing conversion logic
async function handleBatchConvert(uri: vscode.Uri, fileType: string, context: vscode.ExtensionContext) {
    if (!uri) {
        vscode.window.showErrorMessage('Please select a directory for batch conversion.');
        return;
    }

    // Set the file type filter via environment variable
    const originalEnv = process.env.BATCH_FILTER_TYPE;
    process.env.BATCH_FILTER_TYPE = fileType;

    try {
        // Reuse the existing office-to-md conversion logic
        await handleConvertCommand(uri, 'office-to-md', context);
    } finally {
        // Restore the environment variable
        if (originalEnv !== undefined) {
            process.env.BATCH_FILTER_TYPE = originalEnv;
        } else {
            delete process.env.BATCH_FILTER_TYPE;
        }
    }
}

import * as path from 'path';
import * as child_process from 'child_process';

// Batch PDF → PNG handler
async function handleBatchPdfToPng(uri: vscode.Uri, context: vscode.ExtensionContext) {
    if (!uri) {
        vscode.window.showErrorMessage('Please select a directory for batch conversion.');
        return;
    }

    const scriptPath = path.join(context.extensionPath, 'backend', 'converters', 'batch_pdf_to_png.py');
    const targetDir = uri.fsPath;

    // Get the configured python path
    const config = vscode.workspace.getConfiguration('markdownHub');
    const pythonPath = config.get<string>('pythonPath') || 'python';

    // Get the configured poppler path
    let popplerPath = config.get<string>('popplerPath') || '';

    // If not configured, try the bundled poppler
    if (!popplerPath) {
        const localPopplerPath = path.join(context.extensionPath, 'tools', 'poppler', 'poppler-24.02.0', 'Library', 'bin');
        // We can't easily check folder existence without importing fs, so just pass it
        // through; the script will validate.
        popplerPath = localPopplerPath;
    }

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: "Converting single-page PDFs to PNG...",
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
                // Simple progress feedback
                if (msg.includes('Converting')) {
                    progress.report({ message: msg.trim() });
                }
            });

            process.stderr.on('data', (data) => {
                errorOutput += data.toString();
            });

            process.on('close', (code) => {
                if (code === 0) {
                    vscode.window.showInformationMessage(`Batch conversion complete. See the output channel for details.`);
                    resolve();
                } else {
                    vscode.window.showErrorMessage(`Conversion failed (code ${code}): ${errorOutput || output}`);
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
        
        // Register batch conversion commands
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
            (uri: vscode.Uri) => handleBatchPdfToPng(uri, context)),

        vscode.commands.registerCommand('markdown-hub.checkDependencies',
            () => handleCheckDependenciesCommand()),

        vscode.commands.registerCommand('markdown-hub.viewHistory',
            () => handleViewHistoryCommand()),

        vscode.commands.registerCommand('markdown-hub.clearHistory',
            () => handleClearHistoryCommand())
    ];
    
    context.subscriptions.push(...disposables);
}

export function deactivate() {
    console.log('Markdown Hub is now deactivated.');
}