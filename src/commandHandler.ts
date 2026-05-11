import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { executePythonScript } from './pythonService';
import { checkDependencies, checkDependenciesWithQuickPick, DependencyStatus } from './dependencyChecker';

type ConversionType = 'md-to-docx' | 'md-to-pdf' | 'md-to-html' | 'md-to-pptx' | 'office-to-md' | 'diagram-to-png';

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
        console.error('加载历史记录失败:', error);
    }
    return [];
}

function saveHistory(history: HistoryRecord[]): void {
    try {
        const historyData = history.slice(0, MAX_HISTORY_RECORDS);
        fs.writeFileSync(HISTORY_FILE, JSON.stringify(historyData, null, 2), 'utf8');
    } catch (error) {
        console.error('保存历史记录失败:', error);
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
    if (ms < 1000) return `${ms}毫秒`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}秒`;
    return `${(ms / 60000).toFixed(1)}分钟`;
}

function formatTimestamp(timestamp: string): string {
    const date = new Date(timestamp);
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (days === 0) {
        return `今天 ${date.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })}`;
    } else if (days === 1) {
        return `昨天 ${date.toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })}`;
    } else if (days < 7) {
        return `${days}天前`;
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
        'office-to-md': 'Office → Markdown',
        'diagram-to-png': '图表 → PNG'
    };
    return labels[type] || type;
}

function classifyError(errorMessage: string): { category: string; suggestion: string } {
    const lowerError = errorMessage.toLowerCase();

    if (lowerError.includes('permission') || lowerError.includes('权限')) {
        return {
            category: '权限问题',
            suggestion: '请检查文件权限，或以管理员身份运行 VS Code'
        };
    }
    if (lowerError.includes('not found') || lowerError.includes('找不到') || lowerError.includes('不存在')) {
        return {
            category: '文件缺失',
            suggestion: '请确认文件路径正确，且文件未被移动或删除'
        };
    }
    if (lowerError.includes('python') || lowerError.includes('python3')) {
        return {
            category: '依赖缺失',
            suggestion: '请运行"Markdown Hub: Check Dependencies"检查 Python 环境'
        };
    }
    if (lowerError.includes('pandoc') || lowerError.includes('wkhtmltopdf')) {
        return {
            category: '依赖缺失',
            suggestion: '请安装必要的转换工具：pandoc、wkhtmltopdf 等'
        };
    }
    if (lowerError.includes('memory') || lowerError.includes('内存')) {
        return {
            category: '资源不足',
            suggestion: '文件可能过大，建议分批处理或增加系统内存'
        };
    }
    if (lowerError.includes('format') || lowerError.includes('格式')) {
        return {
            category: '格式错误',
            suggestion: '请检查文件格式是否正确，或尝试转换为标准 Markdown 格式'
        };
    }

    return {
        category: '未知错误',
        suggestion: '请查看详细错误信息，或尝试重启 VS Code'
    };
}

function createOutputChannel(): vscode.OutputChannel {
    const channel = vscode.window.createOutputChannel('Markdown Hub');
    return channel;
}

async function handleViewHistoryCommand() {
    const history = loadHistory();

    if (history.length === 0) {
        vscode.window.showInformationMessage('暂无转换历史记录');
        return;
    }

    const channel = createOutputChannel();
    channel.clear();
    channel.appendLine('📋 Markdown Hub - 转换历史记录');
    channel.appendLine('═'.repeat(60));
    channel.appendLine(`共 ${history.length} 条记录 (显示最近 ${Math.min(history.length, MAX_HISTORY_RECORDS)} 条)`);
    channel.appendLine('═'.repeat(60));

    history.slice(0, 20).forEach((record, index) => {
        const statusIcon = record.status === 'success' ? '✅' : '❌';
        const fileName = record.fileName.length > 40 ? record.fileName.substring(0, 37) + '...' : record.fileName;
        const duration = formatDuration(record.duration);
        const timestamp = formatTimestamp(record.timestamp);

        channel.appendLine(`\n${index + 1}. ${statusIcon} ${fileName}`);
        channel.appendLine(`   类型: ${getConversionTypeLabel(record.conversionType)}`);
        channel.appendLine(`   时间: ${timestamp} | 耗时: ${duration}`);
        channel.appendLine(`   状态: ${record.status === 'success' ? '成功' : '失败'}`);

        if (record.status === 'success' && record.outputPath) {
            const outputFileName = path.basename(record.outputPath);
            channel.appendLine(`   输出: ${outputFileName}`);
            if (record.fileSize) {
                channel.appendLine(`   大小: ${formatFileSize(record.fileSize)}`);
            }
        }

        if (record.status === 'failed' && record.errorMessage) {
            const errorPreview = record.errorMessage.length > 50
                ? record.errorMessage.substring(0, 47) + '...'
                : record.errorMessage;
            channel.appendLine(`   错误: ${errorPreview}`);
        }
    });

    if (history.length > 20) {
        channel.appendLine(`\n${'─'.repeat(60)}`);
        channel.appendLine(`还有 ${history.length - 20} 条更早的记录...`);
    }

    channel.show(true);
}

async function handleClearHistoryCommand() {
    const response = await vscode.window.showWarningMessage(
        '确定要清除所有转换历史记录吗？此操作不可撤销。',
        { modal: true },
        '确定清除',
        '取消'
    );

    if (response === '确定清除') {
        try {
            if (fs.existsSync(HISTORY_FILE)) {
                fs.unlinkSync(HISTORY_FILE);
            }
            vscode.window.showInformationMessage('历史记录已清除');
        } catch (error) {
            vscode.window.showErrorMessage('清除历史记录失败');
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
        // 忽略错误
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
 * 处理所有转换命令的核心逻辑
 */
export async function handleConvertCommand(
    resourceUri: vscode.Uri,
    conversionType: ConversionType,
    context: vscode.ExtensionContext
) {
    if (!resourceUri) {
        vscode.window.showErrorMessage('无法执行转换：未选择文件或文件夹。');
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
        channel.appendLine('📁 Markdown Hub - 批量转换');
        channel.appendLine('═'.repeat(60));
        channel.appendLine(`📂 源目录: ${sourcePath}`);
        channel.appendLine(`📋 转换类型: ${getConversionTypeLabel(conversionType)}`);
        channel.appendLine(`📦 待处理: ${totalFiles} 个文件`);
        channel.appendLine('═'.repeat(60) + '\n');
    } else {
        channel.appendLine('🔄 Markdown Hub - 转换开始');
        channel.appendLine('═'.repeat(60));
        channel.appendLine(`📄 文件: ${sourceFileName}`);
        channel.appendLine(`📋 类型: ${getConversionTypeLabel(conversionType)}`);
        channel.appendLine(`💾 大小: ${formatFileSize(fileStats.size)}`);

        if (fileStats.isLarge) {
            channel.appendLine(`⚠️  提示: 大文件转换可能需要较长时间，请耐心等待...\n`);
        } else {
            channel.appendLine('');
        }
    }

    channel.show(true);

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: `正在转换 ${sourceFileName}`,
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

            if (['md-to-docx', 'md-to-pdf', 'md-to-pptx'].includes(conversionType)) {
                const sharedOptions = {
                    projectName: config.get<string>('projectName', ''),
                    author: config.get<string>('author', ''),
                    email: config.get<string>('email', ''),
                    mobilephone: config.get<string>('mobilephone', ''),
                    promoteHeadings: config.get<boolean>('promoteHeadings', true)
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
                        const progressMsg = `正在处理 ${currentFile}/${stats.totalFiles} 个文件`;
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
                    channel.appendLine('✅ 转换成功！');
                    channel.appendLine('═'.repeat(60));
                    channel.appendLine(`⏱️  耗时: ${formatDuration(duration)}`);

                    if (isDir) {
                        channel.appendLine(`📁 已处理: ${totalFiles} 个文件`);
                    }

                    const outputFileName = path.basename(outputFiles[0]);
                    const outputFilePath = outputFiles.length === 1
                        ? outputFiles[0]
                        : path.join(outputDir, `转换结果 (${outputFiles.length}个文件)`);

                    channel.appendLine(`📄 输出: ${outputFileName}`);
                    if (outputFiles.length === 1) {
                        try {
                            const outputStats = fs.statSync(outputFiles[0]);
                            channel.appendLine(`💾 大小: ${formatFileSize(outputStats.size)}`);
                        } catch {
                            // 忽略
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
                        ? `✅ 转换完成！耗时 ${formatDuration(duration)}，输出：${outputFileName}`
                        : `✅ 成功转换 ${outputFiles.length} 个文件！耗时 ${formatDuration(duration)}`;

                    vscode.window.showInformationMessage(message, '打开文件夹').then(selection => {
                        if (selection === '打开文件夹') {
                            vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(outputDir));
                        }
                    });
                }
            } else {
                throw new Error(result.error || '转换失败，发生未知错误');
            }
        } catch (error: any) {
            const endTime = Date.now();
            const duration = endTime - startTime;
            const errorMessage = (error.message || error).toString();
            const errorInfo = classifyError(errorMessage);

            channel.appendLine('');
            channel.appendLine('═'.repeat(60));
            channel.appendLine('❌ 转换失败');
            channel.appendLine('═'.repeat(60));
            channel.appendLine(`⚠️  错误类型: ${errorInfo.category}`);
            channel.appendLine(`💡 建议: ${errorInfo.suggestion}`);
            channel.appendLine(`\n📋 详细错误:\n   ${errorMessage}`);
            channel.appendLine(`⏱️  已耗时: ${formatDuration(duration)}`);

            addHistoryRecord({
                id: `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                fileName: sourceFileName,
                conversionType,
                timestamp: new Date().toISOString(),
                duration,
                status: 'failed',
                errorMessage
            });

            const fullErrorMessage = `${errorInfo.category}：${errorMessage}`;
            vscode.window.showErrorMessage(fullErrorMessage, '查看详情').then(selection => {
                if (selection === '查看详情') {
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
