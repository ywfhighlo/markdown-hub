import { execSync } from 'child_process';
import * as vscode from 'vscode';

export interface DependencyStatus {
    python: boolean;
    pythonVersion?: string;
    pandoc: boolean;
    pandocVersion?: string;
    tesseract: boolean;
    tesseractVersion?: string;
    java: boolean;
    javaVersion?: string;
    drawio: boolean;
    drawioPath?: string;
    mermaidCli: boolean;
    pythonLibs: {
        PyMuPDF: boolean;
        pypdf: boolean;
        pytesseract: boolean;
        psutil: boolean;
        docx2txt: boolean;
        pandas: boolean;
        pythonPptx: boolean;
        html2text: boolean;
    };
}

export interface DependencyIssue {
    name: string;
    severity: 'error' | 'warning' | 'info';
    installCommand?: string;
    description: string;
}

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

export async function checkDependencies(): Promise<DependencyStatus> {
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    const pythonLibNames: Array<keyof DependencyStatus['pythonLibs']> = [
        'PyMuPDF', 'pypdf', 'pytesseract', 'psutil',
        'docx2txt', 'pandas', 'pythonPptx', 'html2text'
    ];

    const pythonLibs: DependencyStatus['pythonLibs'] = {} as any;
    const pythonAvailable = checkCommandExists(pythonCmd);

    if (pythonAvailable) {
        for (const lib of pythonLibNames) {
            pythonLibs[lib] = checkPythonLib(pythonCmd, lib);
        }
    } else {
        for (const lib of pythonLibNames) {
            pythonLibs[lib] = false;
        }
    }

    return {
        python: pythonAvailable,
        pythonVersion: pythonAvailable ? getCommandVersion(pythonCmd, '--version') : undefined,
        pandoc: checkCommandExists('pandoc'),
        pandocVersion: getCommandVersion('pandoc', '--version'),
        tesseract: checkCommandExists('tesseract'),
        tesseractVersion: getCommandVersion('tesseract', '--version'),
        java: checkCommandExists('java'),
        javaVersion: getCommandVersion('java', '-version'),
        drawio: checkCommandExists('drawio') || checkCommandExists('draw.io'),
        mermaidCli: checkCommandExists('mmdc'),
        pythonLibs
    };
}

export async function checkDependenciesWithQuickPick(): Promise<void> {
    const status = await checkDependencies();
    const issues: DependencyIssue[] = [];
    const okItems: string[] = [];

    if (status.python) {
        okItems.push(`✅ Python: ${status.pythonVersion}`);
    } else {
        issues.push({
            name: 'Python',
            severity: 'error',
            installCommand: 'https://www.python.org/downloads/',
            description: 'Python 环境未安装，这是核心依赖'
        });
    }

    if (status.pandoc) {
        okItems.push(`✅ Pandoc: ${status.pandocVersion}`);
    } else {
        issues.push({
            name: 'Pandoc',
            severity: 'error',
            installCommand: process.platform === 'darwin' ? 'brew install pandoc' : 'https://pandoc.org/installing.html',
            description: 'Pandoc 未安装，DOCX/PPTX 转换功能受限'
        });
    }

    if (status.tesseract) {
        okItems.push(`✅ Tesseract OCR`);
    } else {
        issues.push({
            name: 'Tesseract OCR',
            severity: 'warning',
            installCommand: process.platform === 'darwin' ? 'brew install tesseract' : 'https://github.com/UB-Mannheim/tesseract/wiki',
            description: 'Tesseract 未安装，扫描版 PDF OCR 功能不可用'
        });
    }

    if (status.java) {
        okItems.push(`✅ Java: ${status.javaVersion?.split('\n')[0]}`);
    } else {
        issues.push({
            name: 'Java',
            severity: 'warning',
            installCommand: process.platform === 'darwin' ? 'brew install openjdk' : 'https://adoptium.net/',
            description: 'Java 未安装，PlantUML/SVG 转换功能受限'
        });
    }

    if (status.drawio) {
        okItems.push(`✅ draw.io`);
    } else {
        issues.push({
            name: 'draw.io',
            severity: 'info',
            installCommand: process.platform === 'darwin' ? 'brew install --cask drawio' : 'https://github.com/jgraph/drawio-desktop/releases',
            description: 'draw.io 未安装，Draw.io 文件转换功能不可用'
        });
    }

    if (status.mermaidCli) {
        okItems.push(`✅ Mermaid CLI`);
    } else {
        issues.push({
            name: 'Mermaid CLI',
            severity: 'info',
            installCommand: 'npm install -g @mermaid-js/mermaid-cli',
            description: 'Mermaid CLI 未安装，Mermaid 图表转换功能不可用'
        });
    }

    const libIssues = [];
    if (!status.pythonLibs.PyMuPDF) {
        libIssues.push('PyMuPDF');
    }
    if (!status.pythonLibs.pypdf) {
        libIssues.push('pypdf');
    }
    if (!status.pythonLibs.psutil) {
        libIssues.push('psutil');
    }
    if (!status.pythonLibs.docx2txt) {
        libIssues.push('docx2txt');
    }
    if (!status.pythonLibs.pandas) {
        libIssues.push('pandas');
    }
    if (!status.pythonLibs.pythonPptx) {
        libIssues.push('python-pptx');
    }
    if (!status.pythonLibs.html2text) {
        libIssues.push('html2text');
    }

    if (libIssues.length > 0) {
        issues.push({
            name: 'Python 库',
            severity: 'warning',
            installCommand: 'cd backend && pip install -r requirements.txt',
            description: `以下 Python 库缺失: ${libIssues.join(', ')}`
        });
    } else {
        okItems.push(`✅ 所有 Python 库已安装`);
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
