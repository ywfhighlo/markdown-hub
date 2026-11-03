import { execSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vscode from 'vscode';

/**
 * Shared Python discovery for dependencyChecker / pythonService / extension.
 *
 * Resolution order (first verified Python >= 3.8 wins):
 *   1. the `markdown-hub.pythonPath` setting (path or command name)
 *   2. `python` on PATH (Windows) / `python3` on PATH (unix)
 *   3. the `py` launcher on Windows
 *   4. a scan of common install locations:
 *      python.org per-user & all-users, conda, Microsoft Store, py launcher
 *
 * Step 4 exists because "Python is installed but not on PATH" is by far the
 * most common failure on Windows (installer option "Add python.exe to PATH"
 * unchecked), and it makes the extension work out of the box in that case.
 */

export interface PythonCandidate {
    /** Spawnable command: bare name ('python', 'py') or absolute path. */
    cmd: string;
    /** e.g. "Python 3.12.10" (from `--version`); undefined when unverifiable. */
    version?: string;
    source: 'setting' | 'PATH' | 'py-launcher' | 'scan';
}

/** Minimum Python version required by the backend (Python 3.8+). */
const MIN_PY: [number, number] = [3, 8];

/** Quote a command/path when it contains spaces (for cmd.exe / execSync). */
export function q(cmd: string): string {
    return cmd.includes(' ') ? `"${cmd}"` : cmd;
}

function commandVersion(command: string, timeoutMs = 10000): string | undefined {
    try {
        const output = execSync(`${q(command)} --version`, { encoding: 'utf8', timeout: timeoutMs });
        return output.trim().split('\n')[0];
    } catch {
        return undefined;
    }
}

/** "Python 3.12.10" → [3, 12]; undefined when unparsable. */
function parseMajorMinor(version?: string): [number, number] | undefined {
    const match = version?.match(/(\d+)\.(\d+)/);
    if (!match) {
        return undefined;
    }
    return [parseInt(match[1], 10), parseInt(match[2], 10)];
}

function satisfiesMin(version?: string): boolean {
    const mm = parseMajorMinor(version);
    if (!mm) {
        return false;
    }
    return mm[0] > MIN_PY[0] || (mm[0] === MIN_PY[0] && mm[1] >= MIN_PY[1]);
}

/** Read the user-configured `markdown-hub.pythonPath`, if set. */
export function getConfiguredPythonPath(): string | undefined {
    const configured = vscode.workspace.getConfiguration('markdown-hub').get<string>('pythonPath');
    return configured && configured.trim() ? configured.trim() : undefined;
}

/** Numeric version encoded in a name like "Python312" → 312, for desc sorting. */
function numericVersionOf(name: string): number {
    const digits = name.match(/\d+/g);
    return digits ? parseInt(digits.join(''), 10) : 0;
}

/**
 * Scan the most common Python install locations on this machine.
 * Returns existing executables, best-guess order: python.org (newest first,
 * per-user before all-users), conda, py launcher, Microsoft Store.
 * Unverified — callers should confirm with `--version`.
 */
export function scanCommonPythonPaths(): string[] {
    const home = os.homedir();
    const results: string[] = [];

    const pushIfExists = (p: string) => {
        try {
            if (fs.existsSync(p)) {
                results.push(p);
            }
        } catch { /* unreadable — skip */ }
    };

    /** All subdirs of `root` matching `matcher`, newest version first. */
    const sortedSubdirs = (root: string, matcher: (name: string) => boolean): string[] => {
        try {
            if (!fs.existsSync(root)) {
                return [];
            }
            return fs.readdirSync(root)
                .filter(matcher)
                .sort((a, b) => numericVersionOf(b) - numericVersionOf(a));
        } catch {
            return [];
        }
    };

    if (process.platform === 'win32') {
        const localAppData = process.env['LOCALAPPDATA'] || path.join(home, 'AppData', 'Local');
        const programFiles = process.env['ProgramFiles'] || 'C:\\Program Files';
        const programFilesX86 = process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';
        const pythonOrg = (name: string): boolean => /^Python\d+$/.test(name);

        // 1. python.org installs — per-user first (the most common miss-PATH case),
        //    e.g. %LOCALAPPDATA%\Programs\Python\Python312\python.exe
        for (const name of sortedSubdirs(path.join(localAppData, 'Programs', 'Python'), pythonOrg)) {
            pushIfExists(path.join(localAppData, 'Programs', 'Python', name, 'python.exe'));
        }
        for (const name of sortedSubdirs(programFiles, pythonOrg)) {
            pushIfExists(path.join(programFiles, name, 'python.exe'));
        }
        for (const name of sortedSubdirs(programFilesX86, pythonOrg)) {
            pushIfExists(path.join(programFilesX86, name, 'python.exe'));
        }

        // 2. conda (default install locations)
        pushIfExists(path.join(home, 'miniconda3', 'python.exe'));
        pushIfExists(path.join(home, 'anaconda3', 'python.exe'));
        pushIfExists(path.join(localAppData, 'miniconda3', 'python.exe'));
        pushIfExists(path.join(localAppData, 'anaconda3', 'python.exe'));

        // 3. py launcher (installed alongside python.org, but rarely on PATH)
        pushIfExists(path.join(localAppData, 'Programs', 'Python', 'Launcher', 'py.exe'));

        // 4. Microsoft Store real installs: WindowsApps\PythonSoftwareFoundation.Python.3.12_...\python3.12.exe
        //    (the 0-byte python3.exe alias stub in WindowsApps root is deliberately ignored)
        const storeRoot = path.join(localAppData, 'Microsoft', 'WindowsApps');
        for (const name of sortedSubdirs(storeRoot, n => /^PythonSoftwareFoundation/.test(n))) {
            const pkgDir = path.join(storeRoot, name);
            try {
                for (const file of fs.readdirSync(pkgDir)) {
                    if (/^python3(\.\d+)?\.exe$/.test(file)) {
                        results.push(path.join(pkgDir, file));
                    }
                }
            } catch { /* unreadable — skip */ }
        }
    } else {
        // Unix: list real python3.x binaries straight from the bin dirs
        const binDirs = ['/usr/bin', '/usr/local/bin', '/opt/homebrew/bin', '/opt/local/bin'];
        for (const bin of binDirs) {
            try {
                if (!fs.existsSync(bin)) {
                    continue;
                }
                const names = fs.readdirSync(bin)
                    .filter(n => /^python3(\.\d+)?$/.test(n))
                    .sort((a, b) => numericVersionOf(b) - numericVersionOf(a));
                for (const n of names) {
                    results.push(path.join(bin, n));
                }
            } catch { /* unreadable — skip */ }
        }
        pushIfExists(path.join(home, 'miniconda3', 'bin', 'python3'));
        pushIfExists(path.join(home, 'anaconda3', 'bin', 'python3'));
        pushIfExists(path.join(home, '.pyenv', 'shims', 'python3'));
    }

    return [...new Set(results)];
}

/**
 * Ordered candidate list (unverified). First *verified* entry wins in
 * {@link resolvePythonCommand}; broken entries fail their `--version` check
 * and are skipped.
 */
export function pythonCandidates(): PythonCandidate[] {
    const candidates: PythonCandidate[] = [];

    const configured = getConfiguredPythonPath();
    if (configured) {
        candidates.push({ cmd: configured, source: 'setting' });
    }

    if (process.platform === 'win32') {
        candidates.push({ cmd: 'python', source: 'PATH' }, { cmd: 'py', source: 'py-launcher' });
    } else {
        candidates.push({ cmd: 'python3', source: 'PATH' }, { cmd: 'python', source: 'PATH' });
    }

    for (const p of scanCommonPythonPaths()) {
        candidates.push({ cmd: p, source: 'scan' });
    }
    return candidates;
}

/** First candidate whose `--version` works and satisfies Python >= 3.8. */
export function resolvePythonCommand(): PythonCandidate | undefined {
    for (const candidate of pythonCandidates()) {
        const version = commandVersion(candidate.cmd);
        if (satisfiesMin(version)) {
            return { ...candidate, version };
        }
    }
    return undefined;
}
