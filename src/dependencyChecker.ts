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

// Map VS Code conversion types to feature keys
const CONVERSION_TO_FEATURE: Record<string, string[]> = {
    'md-to-docx': ['md_to_docx'],
    'md-to-pdf':  ['md_to_pdf'],
    'md-to-html': ['md_to_html'],
    'md-to-pptx': ['md_to_pptx'],
    'md-to-epub': ['md_to_pdf'],  // EPUB needs pandoc + markdown like PDF
    'office-to-md': ['pdf_to_md', 'word_to_md', 'excel_to_md', 'pptx_to_md', 'html_to_md'],
    'html-to-md': ['html_to_md'],
    'diagram-to-png': ['diagram_to_png'],
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

    // Build a lookup from item label → install command for click-to-install
    const labelToInstall: Record<string, string> = {};
    for (const issue of issues) {
        if (issue.installCommand) {
            labelToInstall[`${issue.name}`] = issue.installCommand;
        }
    }

    const selection = await vscode.window.showQuickPick(items, {
        placeHolder: statusMsg,
        canPickMany: false
    });

    // If user selected an issue item with an install command, run it
    if (selection && selection.label) {
        // Strip icon prefix (❌ / ⚠️ / ℹ️) to match the lookup key
        const label = selection.label.replace(/^[❌⚠️ℹ️✅]\s*/, '');
        const cmd = labelToInstall[label];
        if (cmd) {
            // Check if it's a pip command (can auto-run) vs a URL (open in browser)
            if (cmd.startsWith('pip ') || cmd.startsWith('pip3 ')) {
                const terminal = vscode.window.createTerminal('Markdown Hub — Installing');
                terminal.show();
                terminal.sendText(cmd);
                vscode.window.showInformationMessage(
                    `Installing ${label} dependencies in terminal. Re-check after installation completes.`
                );
            } else if (cmd.startsWith('http')) {
                vscode.env.openExternal(vscode.Uri.parse(cmd));
            } else {
                // Mixed command (pip + URL) — open terminal for pip part
                const pipPart = cmd.split('&&').find(p => p.trim().startsWith('pip'));
                if (pipPart) {
                    const terminal = vscode.window.createTerminal('Markdown Hub — Installing');
                    terminal.show();
                    terminal.sendText(pipPart.trim());
                }
            }
        }
    }
}

/**
 * Pre-check dependencies for a specific conversion type.
 * Returns null if all deps are met, or an object with missing info + install command.
 * This is called BEFORE launching the Python subprocess so the user gets
 * an actionable warning instead of a cryptic post-hoc error.
 */
export async function precheckConversion(
    conversionType: string
): Promise<{ missingLibs: string[]; missingCmds: string[]; installCommand: string; featureName: string } | null> {
    const featureKeys = CONVERSION_TO_FEATURE[conversionType];
    if (!featureKeys) {
        return null;  // Unknown conversion type — skip precheck
    }

    const status = await checkDependencies();
    if (!status.python) {
        return {
            missingLibs: [],
            missingCmds: [],
            installCommand: '',
            featureName: 'Python',
        };
    }

    const allMissingLibs: string[] = [];
    const allMissingCmds: string[] = [];
    const featureNames: string[] = [];

    for (const key of featureKeys) {
        const feat = status.features[key];
        if (feat && !feat.available) {
            allMissingLibs.push(...feat.missingLibs);
            allMissingCmds.push(...feat.missingCmds);
            featureNames.push(feat.name);
        }
    }

    if (allMissingLibs.length === 0 && allMissingCmds.length === 0) {
        return null;  // All good!
    }

    // Deduplicate
    const uniqueLibs = [...new Set(allMissingLibs)];
    const uniqueCmds = [...new Set(allMissingCmds)];

    const installParts: string[] = [];
    if (uniqueLibs.length > 0) {
        installParts.push(pipInstallCmd(uniqueLibs));
    }
    if (uniqueCmds.length > 0) {
        installParts.push(uniqueCmds.map(cmd => installHintFor(cmd)).join('; '));
    }

    return {
        missingLibs: uniqueLibs,
        missingCmds: uniqueCmds,
        installCommand: installParts.join(' && '),
        featureName: featureNames.join(' / ') || conversionType,
    };
}

export { checkDependencies as default };
