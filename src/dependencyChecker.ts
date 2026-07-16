import { execSync } from 'child_process';
import * as vscode from 'vscode';

// ─────────────────────────────────────────
// Per-feature dependency matrix
// Each feature is checked independently: a missing dependency
// only affects its own feature, never the others.
// ─────────────────────────────────────────

/** Dependency status for a single feature. */
export interface FeatureDependency {
    name: string;               // Feature display name
    available: boolean;         // Whether the feature is usable
    missingLibs: string[];      // Missing Python libraries
    missingCmds: string[];      // Missing external commands
    installHint?: string;       // One-line install hint
}

/** Full dependency snapshot. */
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
// Feature definitions: which Python libs and external
// commands each feature requires.
// ─────────────────────────────────────────

interface FeatureDef {
    name: string;
    pythonLibs: string[];       // pip package names (matched against `python -c "import xxx"`)
    commands?: string[];        // Required external commands
    core?: boolean;             // Marks a core feature (missing → severity=error)
}

const FEATURE_DEFS: Record<string, FeatureDef> = {
    'pdf_to_md': {
        name: 'PDF → Markdown',
        pythonLibs: ['PyMuPDF'],   // Minimal core dep; pypdf+pytesseract+pdf2image are optional OCR fallback
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
        name: 'Diagram → PNG',
        pythonLibs: ['Pillow'],
    },
};

// ─────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────

/** Three-platform install hints, kept in sync with backend dep_check._CMD_INSTALL_HINTS. */
const CMD_INSTALL_HINTS: Record<string, Record<string, string>> = {
    pandoc: {
        win32:  '下载 https://pandoc.org/installing.html',
        darwin: 'brew install pandoc',
        linux:  'sudo apt install pandoc   # 或 dnf install pandoc',
    },
    tesseract: {
        win32:  '下载 https://github.com/UB-Mannheim/tesseract/wiki',
        darwin: 'brew install tesseract',
        linux:  'sudo apt install tesseract-ocr',
    },
    java: {
        win32:  '下载 https://adoptium.net/',
        darwin: 'brew install openjdk',
        linux:  'sudo apt install openjdk-11-jdk',
    },
    graphviz: {
        win32:  '下载 https://graphviz.org/download/ 并将 bin 加入 PATH',
        darwin: 'brew install graphviz',
        linux:  'sudo apt install graphviz   # 或 dnf install graphviz',
    },
    poppler: {
        win32:  '首用时自动下载；或手动下载 https://github.com/oschwartz10612/poppler-windows/releases',
        darwin: 'brew install poppler',
        linux:  'sudo apt install poppler-utils',
    },
    drawio: {
        win32:  '下载 https://github.com/jgraph/drawio-desktop/releases',
        darwin: 'brew install --cask drawio',
        linux:  '下载 https://github.com/jgraph/drawio-desktop/releases',
    },
    mmdc: {
        win32:  'npm install -g @mermaid-js/mermaid-cli',
        darwin: 'npm install -g @mermaid-js/mermaid-cli',
        linux:  'npm install -g @mermaid-js/mermaid-cli',
    },
};

function platformKey(): 'win32' | 'darwin' | 'linux' {
    if (process.platform.startsWith('win')) return 'win32';
    if (process.platform === 'darwin') return 'darwin';
    return 'linux';
}

function installHintFor(cmd: string): string {
    const hints = CMD_INSTALL_HINTS[cmd];
    if (!hints) return cmd;
    return hints[platformKey()] || hints.linux || cmd;
}

function checkCommandExists(command: string, versionFlag: string = '--version'): boolean {
    try {
        execSync(`${command} ${versionFlag}`, { stdio: 'ignore' });
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

/** Build a platform-appropriate `pip install` command. */
function pipInstallCmd(libs: string[]): string {
    return `pip install ${libs.join(' ')}`;
}

// ─────────────────────────────────────────
// Main check logic
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
                ...(missingCmds.length > 0 ? [missingCmds.map(cmd => installHintFor(cmd)).join(', ')] : []),
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
// QuickPick display
// ─────────────────────────────────────────

export async function checkDependenciesWithQuickPick(): Promise<void> {
    const status = await checkDependencies();
    const issues: DependencyIssue[] = [];
    const okItems: string[] = [];

    // Python itself
    if (status.python) {
        okItems.push(`✅ Python: ${status.pythonVersion}`);
    } else {
        issues.push({
            name: 'Python',
            severity: 'error',
            installCommand: 'https://www.python.org/downloads/',
            description: 'Python is not installed — all features are unavailable'
        });
    }

    // Show each feature's status
    for (const [key, feat] of Object.entries(status.features)) {
        const def = FEATURE_DEFS[key];
        if (feat.available) {
            okItems.push(`✅ ${feat.name}`);
        } else {
            const parts: string[] = [];
            if (feat.missingLibs.length > 0) {
                parts.push(`Missing Python libs: ${feat.missingLibs.join(', ')}`);
            }
            if (feat.missingCmds.length > 0) {
                parts.push(`Missing external tools: ${feat.missingCmds.join(', ')}`);
            }

            issues.push({
                name: feat.name,
                severity: def?.core ? 'error' : 'warning',
                installCommand: feat.installHint,
                description: parts.join('；'),
            });
        }
    }

    // Optional external tools (diagram-related; not tied to a specific conversion feature)
    const optionalTools: Array<{ name: string; cmd: string; hint: string; versionFlag?: string }> = [
        { name: 'Tesseract OCR', cmd: 'tesseract', hint: installHintFor('tesseract') },
        { name: 'Java', cmd: 'java', hint: installHintFor('java') },
        { name: 'Graphviz', cmd: 'dot', hint: installHintFor('graphviz'), versionFlag: '-V' },
        { name: 'draw.io', cmd: 'drawio', hint: installHintFor('drawio') },
        { name: 'Mermaid CLI', cmd: 'mmdc', hint: installHintFor('mmdc') },
    ];

    for (const tool of optionalTools) {
        const flag = tool.versionFlag || '--version';
        if (checkCommandExists(tool.cmd, flag) || checkCommandExists(tool.cmd === 'drawio' ? 'draw.io' : tool.cmd, flag)) {
            okItems.push(`✅ ${tool.name}`);
        } else {
            issues.push({
                name: tool.name,
                severity: 'info',
                installCommand: tool.hint,
                description: 'Not installed — some diagram conversions may be unavailable',
            });
        }
    }

    const allOk = issues.length === 0;
    const statusMsg = allOk ? '🎉 All dependencies are ready!' : `${issues.length} issue(s) found`;

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
                description: issue.installCommand ? `💡 Install: ${issue.installCommand}` : undefined
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
