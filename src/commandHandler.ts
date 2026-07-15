import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { executePythonScript } from './pythonService';
import { checkDependencies, checkDependenciesWithQuickPick, DependencyStatus } from './dependencyChecker';

type ConversionType = 'md-to-docx' | 'md-to-pdf' | 'md-to-html' | 'md-to-pptx' | 'md-to-epub' | 'office-to-md' | 'diagram-to-png' | 'html-to-md';

interface HistoryRecord {
    id: string;
    fileName: string;
    conversionType: ConversionType;
    timestamp: string;
    duration: number;
    status: 'success' | 'failed';
    outputPath?: string;
    fileSize?: number;
    errorMessage?: string;
}

interface ConversionStats {
    totalFiles?: number;
    currentFile?: number;
    fileSize?: number;
    pageCount?: number;
}

const HISTORY_FILE = path.join(os.homedir(), '.markdown-hub', 'history.json');
const MAX_HISTORY_RECORDS = 50;

function getHistoryFilePath(): string {
    const dir = path.dirname(HISTORY_FILE);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    return HISTORY_FILE;
}

function loadHistory(): HistoryRecord[] {
    try {
        if (fs.existsSync(HISTORY_FILE)) {
            const data = fs.readFileSync(HISTORY_FILE, 'utf8');
            return JSON.parse(data);
        }
    } catch (error) {
        console.error('Failed to load history:', error);
    }
    return [];
}

function saveHistory(history: HistoryRecord[]): void {
    try {
        const historyData = history.slice(0, MAX_HISTORY_RECORDS);
        fs.writeFileSync(HISTORY_FILE, JSON.stringify(historyData, null, 2), 'utf8');
    } catch (error) {
        console.error('Failed to save history:', error);
    }
}

function addHistoryRecord(record: HistoryRecord): void {
    const history = loadHistory();
    history.unshift(record);
    saveHistory(history);
}

function formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDuration(ms: number): string {
    if (ms < 1000) return `${ms}ms`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
    return `${(ms / 60000).toFixed(1)}min`;
}

function formatTimestamp(timestamp: string): string {
    const date = new Date(timestamp);
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (days === 0) {
        return `Today ${date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })}`;
    } else if (days === 1) {
        return `Yesterday ${date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })}`;
    } else if (days < 7) {
        return `${days} days ago`;
    } else {
        return date.toLocaleDateString('zh-CN', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
    }
}

function getConversionTypeLabel(type: ConversionType): string {
    const labels: Record<ConversionType, string> = {
        'md-to-docx': 'Markdown → Word',
        'md-to-pdf': 'Markdown → PDF',
        'md-to-html': 'Markdown → HTML',
        'md-to-pptx': 'Markdown → PPT',
        'md-to-epub': 'Markdown → EPUB',
        'office-to-md': 'Office → Markdown',
        'diagram-to-png': 'Diagram → PNG',
        'html-to-md': 'HTML → Markdown'
    };
    return labels[type] || type;
}

function classifyError(errorMessage: string): { category: string; suggestion: string } {
    const lowerError = errorMessage.toLowerCase();

    if (lowerError.includes('permission') || lowerError.includes('denied')) {
        return {
            category: 'Permission Denied',
            suggestion: 'Check file permissions, or run VS Code as administrator'
        };
    }
    if (lowerError.includes('not found') || lowerError.includes('does not exist') || lowerError.includes('no such file')) {
        return {
            category: 'File Missing',
            suggestion: 'Verify the file path is correct and the file has not been moved or deleted'
        };
    }
    if (lowerError.includes('python') || lowerError.includes('python3')) {
        return {
            category: 'Missing Dependency',
            suggestion: 'Run "Markdown Hub: Check Dependencies" to verify your Python environment'
        };
    }
    if (lowerError.includes('pandoc') || lowerError.includes('wkhtmltopdf')) {
        return {
            category: 'Missing Dependency',
            suggestion: 'Install the required conversion tools: pandoc, wkhtmltopdf, etc.'
        };
    }
    if (lowerError.includes('memory') || lowerError.includes('out of memory')) {
        return {
            category: 'Insufficient Resources',
            suggestion: 'The file may be too large — try batch processing or increase system memory'
        };
    }
    if (lowerError.includes('format')) {
        return {
            category: 'Format Error',
            suggestion: 'Check that the file format is correct, or try converting to standard Markdown'
        };
    }

    return {
        category: 'Unknown Error',
        suggestion: 'See the detailed error output, or try restarting VS Code'
    };
}

function createOutputChannel(): vscode.OutputChannel {
    const channel = vscode.window.createOutputChannel('Markdown Hub');
    return channel;
}

async function handleViewHistoryCommand() {
    const history = loadHistory();

    if (history.length === 0) {
        vscode.window.showInformationMessage('No conversion history yet.');
        return;
    }

    const channel = createOutputChannel();
    channel.clear();
    channel.appendLine('📋 Markdown Hub - Conversion History');
    channel.appendLine('═'.repeat(60));
    channel.appendLine(`${history.length} record(s) total (showing the most recent ${Math.min(history.length, MAX_HISTORY_RECORDS)})`);
    channel.appendLine('═'.repeat(60));

    history.slice(0, 20).forEach((record, index) => {
        const statusIcon = record.status === 'success' ? '✅' : '❌';
        const fileName = record.fileName.length > 40 ? record.fileName.substring(0, 37) + '...' : record.fileName;
        const duration = formatDuration(record.duration);
        const timestamp = formatTimestamp(record.timestamp);

        channel.appendLine(`\n${index + 1}. ${statusIcon} ${fileName}`);
        channel.appendLine(`   Type: ${getConversionTypeLabel(record.conversionType)}`);
        channel.appendLine(`   Time: ${timestamp} | Duration: ${duration}`);
        channel.appendLine(`   Status: ${record.status === 'success' ? 'Success' : 'Failed'}`);

        if (record.status === 'success' && record.outputPath) {
            const outputFileName = path.basename(record.outputPath);
            channel.appendLine(`   Output: ${outputFileName}`);
            if (record.fileSize) {
                channel.appendLine(`   Size: ${formatFileSize(record.fileSize)}`);
            }
        }

        if (record.status === 'failed' && record.errorMessage) {
            const errorPreview = record.errorMessage.length > 50
                ? record.errorMessage.substring(0, 47) + '...'
                : record.errorMessage;
            channel.appendLine(`   Error: ${errorPreview}`);
        }
    });

    if (history.length > 20) {
        channel.appendLine(`\n${'─'.repeat(60)}`);
        channel.appendLine(`${history.length - 20} older record(s) not shown...`);
    }

    channel.show(true);
}

async function handleClearHistoryCommand() {
    const response = await vscode.window.showWarningMessage(
        'Clear all conversion history? This cannot be undone.',
        { modal: true },
        'Clear',
        'Cancel'
    );

    if (response === 'Clear') {
        try {
            if (fs.existsSync(HISTORY_FILE)) {
                fs.unlinkSync(HISTORY_FILE);
            }
            vscode.window.showInformationMessage('History cleared.');
        } catch (error) {
            vscode.window.showErrorMessage('Failed to clear history.');
        }
    }
}

async function handleCheckDependenciesCommand() {
    await checkDependenciesWithQuickPick();
}

function isDirectory(sourcePath: string): boolean {
    try {
        return fs.statSync(sourcePath).isDirectory();
    } catch {
        return false;
    }
}

function countFiles(dirPath: string, extensions: string[]): number {
    let count = 0;
    try {
        const items = fs.readdirSync(dirPath);
        for (const item of items) {
            const fullPath = path.join(dirPath, item);
            if (fs.statSync(fullPath).isDirectory()) {
                count += countFiles(fullPath, extensions);
            } else {
                const ext = path.extname(item).toLowerCase();
                if (extensions.includes(ext)) {
                    count++;
                }
            }
        }
    } catch {
        // ignore error
    }
    return count;
}

function getFileStats(sourcePath: string): { size: number; isLarge: boolean } {
    try {
        const stats = fs.statSync(sourcePath);
        const size = stats.size;
        return {
            size,
            isLarge: size > 10 * 1024 * 1024
        };
    } catch {
        return { size: 0, isLarge: false };
    }
}

/**
 * Core handler for all conversion commands.
 */
export async function handleConvertCommand(
    resourceUri: vscode.Uri,
    conversionType: ConversionType,
    context: vscode.ExtensionContext
) {
    if (!resourceUri) {
        vscode.window.showErrorMessage('Cannot convert: no file or folder selected.');
        return;
    }

    const sourcePath = resourceUri.fsPath;
    const config = vscode.workspace.getConfiguration('markdown-hub');
    const channel = createOutputChannel();

    const startTime = Date.now();
    const sourceFileName = path.basename(sourcePath);
    const isDir = isDirectory(sourcePath);
    const fileStats = getFileStats(sourcePath);

    let totalFiles = 1;
    let currentFile = 0;

    if (isDir) {
        const extensions = conversionType === 'office-to-md'
            ? ['.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls']
            : ['.md', '.markdown'];
        totalFiles = countFiles(sourcePath, extensions);
        channel.appendLine('📁 Markdown Hub - Batch Conversion');
        channel.appendLine('═'.repeat(60));
        channel.appendLine(`📂 Source dir: ${sourcePath}`);
        channel.appendLine(`📋 Type: ${getConversionTypeLabel(conversionType)}`);
        channel.appendLine(`📦 Pending: ${totalFiles} file(s)`);
        channel.appendLine('═'.repeat(60) + '\n');
    } else {
        channel.appendLine('🔄 Markdown Hub - Conversion Started');
        channel.appendLine('═'.repeat(60));
        channel.appendLine(`📄 File: ${sourceFileName}`);
        channel.appendLine(`📋 Type: ${getConversionTypeLabel(conversionType)}`);
        channel.appendLine(`💾 Size: ${formatFileSize(fileStats.size)}`);

        if (fileStats.isLarge) {
            channel.appendLine(`⚠️  Note: large file conversion may take a while, please wait...\n`);
        } else {
            channel.appendLine('');
        }
    }

    channel.show(true);

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: `Converting ${sourceFileName}`,
        cancellable: false
    }, async (progress) => {
        try {
            const configOutputDir = config.get<string>('outputDirectory', './converted');
            const outputDir = path.isAbsolute(configOutputDir)
                ? configOutputDir
                : path.join(path.dirname(sourcePath), configOutputDir);

            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }

            let conversionOptions: any = null;

            if (['md-to-docx', 'md-to-pdf', 'md-to-pptx', 'md-to-epub'].includes(conversionType)) {
                const sharedOptions = {
                    projectName: config.get<string>('projectName', ''),
                    author: config.get<string>('author', ''),
                    email: config.get<string>('email', ''),
                    mobilephone: config.get<string>('mobilephone', ''),
                    promoteHeadings: config.get<boolean>('promoteHeadings', true),
                    // Code block highlight theme (pandoc --highlight-style), only affects DOCX/PDF/HTML
                    codeHighlightTheme: config.get<string>('codeHighlightTheme', 'pygments')
                };
                conversionOptions = { ...sharedOptions };

                if (conversionType === 'md-to-docx' || conversionType === 'md-to-pdf') {
                    if (config.get<boolean>('useDocxTemplate', true)) {
                        let templatePath = config.get<string>('docxTemplatePath', '');
                        if (!templatePath || templatePath.trim() === '') {
                            templatePath = path.join(context.extensionPath, 'backend', 'converters', 'templates', 'template.docx');
                        }
                        conversionOptions.docxTemplatePath = templatePath;
                    }

                    conversionOptions.svgDpi = config.get<number>('svgDpi', 300);
                    conversionOptions.svgOutputWidth = config.get<number>('svgOutputWidth', 800);
                } else if (conversionType === 'md-to-pptx') {
                    if (config.get<boolean>('usePptxTemplate', true)) {
                        let templatePath = config.get<string>('pptxTemplatePath', '');
                        if (!templatePath || templatePath.trim() === '') {
                            templatePath = path.join(context.extensionPath, 'backend', 'converters', 'templates', 'template.pptx');
                        }
                        conversionOptions.pptxTemplatePath = templatePath;
                    }

                    conversionOptions.svgDpi = config.get<number>('svgDpi', 300);
                    conversionOptions.svgOutputWidth = config.get<number>('svgOutputWidth', 800);
                } else if (conversionType === 'md-to-html') {
                    conversionOptions.svgDpi = config.get<number>('svgDpi', 300);
                    conversionOptions.svgOutputWidth = config.get<number>('svgOutputWidth', 800);
                } else if (conversionType === 'md-to-epub') {
                    conversionOptions.svgDpi = config.get<number>('svgDpi', 300);
                    conversionOptions.svgOutputWidth = config.get<number>('svgOutputWidth', 800);
                    conversionOptions.codeHighlightTheme = config.get<string>('codeHighlightTheme', 'pygments');
                }
            } else if (conversionType === 'office-to-md') {
                conversionOptions = {
                    popplerPath: config.get<string>('popplerPath', ''),
                    tesseractCmd: config.get<string>('tesseractCmd', '')
                };
            }

            const result = await executePythonScript(
                sourcePath,
                conversionType,
                outputDir,
                context,
                conversionOptions,
                (message: string, percentage?: number, stats?: ConversionStats) => {
                    if (stats && stats.totalFiles && stats.totalFiles > 1) {
                        currentFile = stats.currentFile || 0;
                        const progressMsg = `Processing ${currentFile}/${stats.totalFiles}`;
                        channel.appendLine(`📊 ${progressMsg}: ${message}`);
                        progress.report({
                            message: progressMsg,
                            increment: percentage !== undefined
                                ? percentage - (progress as any).value || 0
                                : undefined
                        });
                    } else {
                        const progressMsg = message;
                        channel.appendLine(`📊 ${progressMsg}`);
                        progress.report({
                            message: progressMsg,
                            increment: percentage !== undefined
                                ? percentage - (progress as any).value || 0
                                : undefined
                        });
                    }
                }
            );

            const endTime = Date.now();
            const duration = endTime - startTime;

            if (result.success) {
                const outputFiles = result.outputFiles || [];
                if (outputFiles.length > 0) {
                    channel.appendLine('');
                    channel.appendLine('═'.repeat(60));
                    channel.appendLine('✅ Conversion succeeded!');
                    channel.appendLine('═'.repeat(60));
                    channel.appendLine(`⏱️  Duration: ${formatDuration(duration)}`);

                    if (isDir) {
                        channel.appendLine(`📁 Processed: ${totalFiles} file(s)`);
                    }

                    const outputFileName = path.basename(outputFiles[0]);
                    const outputFilePath = outputFiles.length === 1
                        ? outputFiles[0]
                        : path.join(outputDir, `conversion-result (${outputFiles.length} files)`);

                    channel.appendLine(`📄 Output: ${outputFileName}`);
                    if (outputFiles.length === 1) {
                        try {
                            const outputStats = fs.statSync(outputFiles[0]);
                            channel.appendLine(`💾 Size: ${formatFileSize(outputStats.size)}`);
                        } catch {
                            // ignore
                        }
                    }

                    addHistoryRecord({
                        id: `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                        fileName: sourceFileName,
                        conversionType,
                        timestamp: new Date().toISOString(),
                        duration,
                        status: 'success',
                        outputPath: outputFilePath,
                        fileSize: fileStats.size
                    });

                    const message = outputFiles.length === 1
                        ? `✅ Conversion complete! (${formatDuration(duration)}) — ${outputFileName}`
                        : `✅ Converted ${outputFiles.length} file(s)! (${formatDuration(duration)})`;

                    vscode.window.showInformationMessage(message, 'Reveal in Folder').then(selection => {
                        if (selection === 'Reveal in Folder') {
                            vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(outputDir));
                        }
                    });
                }
            } else {
                throw new Error(result.error || 'Conversion failed: unknown error');
            }
        } catch (error: any) {
            const endTime = Date.now();
            const duration = endTime - startTime;
            const errorMessage = (error.message || error).toString();
            const errorInfo = classifyError(errorMessage);

            channel.appendLine('');
            channel.appendLine('═'.repeat(60));
            channel.appendLine('❌ Conversion failed');
            channel.appendLine('═'.repeat(60));
            channel.appendLine(`⚠️  Error type: ${errorInfo.category}`);
            channel.appendLine(`💡 Suggestion: ${errorInfo.suggestion}`);
            channel.appendLine(`\n📋 Detailed error:\n   ${errorMessage}`);
            channel.appendLine(`⏱️  Elapsed: ${formatDuration(duration)}`);

            addHistoryRecord({
                id: `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                fileName: sourceFileName,
                conversionType,
                timestamp: new Date().toISOString(),
                duration,
                status: 'failed',
                errorMessage
            });

            const fullErrorMessage = `${errorInfo.category}: ${errorMessage}`;
            vscode.window.showErrorMessage(fullErrorMessage, 'View Details').then(selection => {
                if (selection === 'View Details') {
                    channel.show(true);
                }
            });
        }
    });
}

export async function handleOpenTemplateSettingsCommand() {
    await vscode.commands.executeCommand('workbench.action.openSettings', '@ext:ywfhighlo.markdown-hub template');
}

export { handleViewHistoryCommand, handleClearHistoryCommand, handleCheckDependenciesCommand };
