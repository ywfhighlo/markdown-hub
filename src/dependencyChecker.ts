import { execSync } from 'child_process';
import * as vscode from 'vscode';

// ─────────────────────────────────────────
// 功能维度依赖矩阵
// 每个功能独立检查，缺失只影响该功能，不波及其他
// ─────────────────────────────────────────

/** 单个功能维度的依赖状态 */
export interface FeatureDependency {
    name: string;               // 功能名（中文）
    available: boolean;         // 是否可用
    missingLibs: string[];      // 缺失的 Python 库
    missingCmds: string[];      // 缺失的外部命令
    installHint?: string;       // 一行安装提示
}

/** 全量依赖快照 */
export interface DependencyStatus {
    python: boolean;
    pythonVersion?: string;
    features: Record<string, FeatureDependency>;
}

export interface DependencyIssue {
    name: string;
    severity: 'error' | 'warning' | 'info';
    installCommand?: string;
    description: string;
}

// ─────────────────────────────────────────
// 功能定义：每个功能需要哪些 Python 库和外部命令
// ─────────────────────────────────────────

interface FeatureDef {
    name: string;
    pythonLibs: string[];       // pip 包名（与 python -c "import xxx" 对应）
    commands?: string[];        // 需要的外部命令
    core?: boolean;             // 标记为核心功能（缺失时 severity=error）
}

const FEATURE_DEFS: Record<string, FeatureDef> = {
    'pdf_to_md': {
        name: 'PDF → Markdown',
        pythonLibs: ['PyMuPDF'],   // 核心最小依赖；pypdf+pytesseract+pdf2image 为可选 OCR 回退
        core: true,
    },
    'word_to_md': {
        name: 'Word → Markdown',
        pythonLibs: ['docx2txt'],
    },
    'excel_to_md': {
        name: 'Excel → Markdown',
        pythonLibs: ['pandas', 'tabulate', 'openpyxl'],
    },
    'pptx_to_md': {
        name: 'PPTX → Markdown',
        pythonLibs: ['python-pptx'],
    },
    'html_to_md': {
        name: 'HTML → Markdown',
        pythonLibs: ['html2text'],
    },
    'md_to_docx': {
        name: 'Markdown → DOCX',
        pythonLibs: ['python-docx', 'docxtpl', 'docxcompose', 'docx2txt'],
        commands: ['pandoc'],
    },
    'md_to_pdf': {
        name: 'Markdown → PDF',
        pythonLibs: ['markdown'],
        commands: ['pandoc'],
    },
    'md_to_html': {
        name: 'Markdown → HTML',
        pythonLibs: ['markdown'],
    },
    'md_to_pptx': {
        name: 'Markdown → PPTX',
        pythonLibs: ['python-pptx', 'Pillow'],
    },
    'diagram_to_png': {
        name: '图表 → PNG',
        pythonLibs: ['Pillow'],
    },
};

// ─────────────────────────────────────────
// 工具函数
// ─────────────────────────────────────────

function checkCommandExists(command: string): boolean {
    try {
        execSync(`${command} --version`, { stdio: 'ignore' });
        return true;
    } catch {
        return false;
    }
}

function getCommandVersion(command: string, versionFlag: string = '--version'): string | undefined {
    try {
        const output = execSync(`${command} ${versionFlag}`, { encoding: 'utf8', timeout: 10000 });
        return output.trim().split('\n')[0];
    } catch {
        return undefined;
    }
}

function checkPythonLib(pythonCmd: string, libName: string): boolean {
    try {
        execSync(`${pythonCmd} -c "import ${libName.replace('-', '_')}"`, { stdio: 'ignore' });
        return true;
    } catch {
        return false;
    }
}

/** 按平台生成 pip install 命令 */
function pipInstallCmd(libs: string[]): string {
    return `pip install ${libs.join(' ')}`;
}

// ─────────────────────────────────────────
// 主检查逻辑
// ─────────────────────────────────────────

export async function checkDependencies(): Promise<DependencyStatus> {
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    const pythonAvailable = checkCommandExists(pythonCmd);

    const features: Record<string, FeatureDependency> = {};

    for (const [key, def] of Object.entries(FEATURE_DEFS)) {
        const missingLibs = pythonAvailable
            ? def.pythonLibs.filter(lib => !checkPythonLib(pythonCmd, lib))
            : [...def.pythonLibs];

        const missingCmds = (def.commands || []).filter(cmd => !checkCommandExists(cmd));

        features[key] = {
            name: def.name,
            available: missingLibs.length === 0 && missingCmds.length === 0,
            missingLibs,
            missingCmds,
            installHint: [
                ...(missingLibs.length > 0 ? [pipInstallCmd(missingLibs)] : []),
                ...(missingCmds.length > 0 ? [missingCmds.map(cmd => {
                    switch (cmd) {
                        case 'pandoc': return process.platform === 'darwin'
                            ? 'brew install pandoc'
                            : 'https://pandoc.org/installing.html';
                        default: return cmd;
                    }
                }).join(', ')] : []),
            ].join(' && ') || undefined,
        };
    }

    return {
        python: pythonAvailable,
        pythonVersion: pythonAvailable ? getCommandVersion(pythonCmd, '--version') : undefined,
        features,
    };
}

// ─────────────────────────────────────────
// QuickPick 展示
// ─────────────────────────────────────────

export async function checkDependenciesWithQuickPick(): Promise<void> {
    const status = await checkDependencies();
    const issues: DependencyIssue[] = [];
    const okItems: string[] = [];

    // Python 本身
    if (status.python) {
        okItems.push(`✅ Python: ${status.pythonVersion}`);
    } else {
        issues.push({
            name: 'Python',
            severity: 'error',
            installCommand: 'https://www.python.org/downloads/',
            description: 'Python 环境未安装，所有功能不可用'
        });
    }

    // 按功能维度逐一展示
    for (const [key, feat] of Object.entries(status.features)) {
        const def = FEATURE_DEFS[key];
        if (feat.available) {
            okItems.push(`✅ ${feat.name}`);
        } else {
            const parts: string[] = [];
            if (feat.missingLibs.length > 0) {
                parts.push(`Python库缺失: ${feat.missingLibs.join(', ')}`);
            }
            if (feat.missingCmds.length > 0) {
                parts.push(`外部工具缺失: ${feat.missingCmds.join(', ')}`);
            }

            issues.push({
                name: feat.name,
                severity: def?.core ? 'error' : 'warning',
                installCommand: feat.installHint,
                description: parts.join('；'),
            });
        }
    }

    // 可选的外部工具（图表相关，不属于某个特定转换功能）
    const optionalTools: Array<{ name: string; cmd: string; hint: string }> = [
        { name: 'Tesseract OCR', cmd: 'tesseract', hint: process.platform === 'darwin' ? 'brew install tesseract' : 'https://github.com/UB-Mannheim/tesseract/wiki' },
        { name: 'Java', cmd: 'java', hint: process.platform === 'darwin' ? 'brew install openjdk' : 'https://adoptium.net/' },
        { name: 'draw.io', cmd: 'drawio', hint: process.platform === 'darwin' ? 'brew install --cask drawio' : 'https://github.com/jgraph/drawio-desktop/releases' },
        { name: 'Mermaid CLI', cmd: 'mmdc', hint: 'npm install -g @mermaid-js/mermaid-cli' },
    ];

    for (const tool of optionalTools) {
        if (checkCommandExists(tool.cmd) || checkCommandExists(tool.cmd === 'drawio' ? 'draw.io' : tool.cmd)) {
            okItems.push(`✅ ${tool.name}`);
        } else {
            issues.push({
                name: tool.name,
                severity: 'info',
                installCommand: tool.hint,
                description: '未安装，部分图表转换功能不可用',
            });
        }
    }

    const allOk = issues.length === 0;
    const statusMsg = allOk ? '🎉 所有依赖已就绪！' : `发现 ${issues.length} 个问题`;

    const items: vscode.QuickPickItem[] = [];

    if (issues.length > 0) {
        items.push({
            label: `⚠️  ${statusMsg}`,
            kind: vscode.QuickPickItemKind.Separator
        } as any);

        for (const issue of issues) {
            const icon = issue.severity === 'error' ? '❌' : issue.severity === 'warning' ? '⚠️' : 'ℹ️';
            items.push({
                label: `${icon} ${issue.name}`,
                detail: issue.description,
                description: issue.installCommand ? `💡 安装: ${issue.installCommand}` : undefined
            });
        }

        items.push({
            label: '',
            kind: vscode.QuickPickItemKind.Separator
        } as any);
    }

    for (const ok of okItems) {
        items.push({
            label: ok
        } as any);
    }

    await vscode.window.showQuickPick(items, {
        placeHolder: statusMsg,
        canPickMany: false
    });
}

export { checkDependencies as default };
