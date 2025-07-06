#!/usr/bin/env node

/**
 * 将 Python 依赖打包到 VSCode 扩展中的脚本
 * 这样用户安装扩展后就不需要单独安装 Python 依赖了
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const BUNDLE_DIR = path.join(__dirname, '..', 'bundled');
const BACKEND_DIR = path.join(__dirname, '..', 'backend');
const PYTHON_DEPS_DIR = path.join(BUNDLE_DIR, 'python-deps');

console.log('🚀 开始打包 Python 依赖...');

// 创建打包目录
if (!fs.existsSync(BUNDLE_DIR)) {
    fs.mkdirSync(BUNDLE_DIR, { recursive: true });
}

if (!fs.existsSync(PYTHON_DEPS_DIR)) {
    fs.mkdirSync(PYTHON_DEPS_DIR, { recursive: true });
}

// 检查 Python 是否可用
function checkPython() {
    const pythonCommands = ['python', 'python3', 'py'];
    
    for (const cmd of pythonCommands) {
        try {
            execSync(`${cmd} --version`, { stdio: 'pipe' });
            console.log(`✓ 找到 Python 命令: ${cmd}`);
            return cmd;
        } catch (e) {
            // 继续尝试下一个命令
        }
    }
    
    throw new Error('未找到 Python 命令，请确保 Python 已安装并在 PATH 中');
}

// 安装依赖到指定目录
function installDependencies(pythonCmd) {
    console.log('📦 安装 Python 依赖到打包目录...');
    
    const requirementsFile = path.join(BACKEND_DIR, 'requirements.txt');
    const minimalRequirementsFile = path.join(BACKEND_DIR, 'requirements-minimal.txt');
    
    // 首先尝试安装完整依赖
    try {
        console.log('尝试安装完整依赖...');
        execSync(`${pythonCmd} -m pip install --target "${PYTHON_DEPS_DIR}" -r "${requirementsFile}"`, {
            stdio: 'inherit'
        });
        console.log('✓ 完整依赖安装成功');
        return true;
    } catch (error) {
        console.log('⚠️ 完整依赖安装失败，尝试最小依赖...');
        
        // 清理失败的安装
        if (fs.existsSync(PYTHON_DEPS_DIR)) {
            fs.rmSync(PYTHON_DEPS_DIR, { recursive: true, force: true });
            fs.mkdirSync(PYTHON_DEPS_DIR, { recursive: true });
        }
        
        try {
            execSync(`${pythonCmd} -m pip install --target "${PYTHON_DEPS_DIR}" -r "${minimalRequirementsFile}"`, {
                stdio: 'inherit'
            });
            console.log('✓ 最小依赖安装成功');
            return false;
        } catch (minimalError) {
            throw new Error('无法安装 Python 依赖，请检查网络连接和 Python 环境');
        }
    }
}

// 复制后端文件
function copyBackendFiles() {
    console.log('📁 复制后端文件...');
    
    const backendTargetDir = path.join(BUNDLE_DIR, 'backend');
    
    // 创建目标目录
    if (fs.existsSync(backendTargetDir)) {
        fs.rmSync(backendTargetDir, { recursive: true, force: true });
    }
    fs.mkdirSync(backendTargetDir, { recursive: true });
    
    // 复制文件
    const filesToCopy = [
        'cli.py',
        'converters',
        'requirements.txt',
        'requirements-minimal.txt'
    ];
    
    filesToCopy.forEach(file => {
        const sourcePath = path.join(BACKEND_DIR, file);
        const targetPath = path.join(backendTargetDir, file);
        
        if (fs.existsSync(sourcePath)) {
            if (fs.statSync(sourcePath).isDirectory()) {
                fs.cpSync(sourcePath, targetPath, { recursive: true });
            } else {
                fs.copyFileSync(sourcePath, targetPath);
            }
            console.log(`✓ 复制: ${file}`);
        } else {
            console.log(`⚠️ 未找到文件: ${file}`);
        }
    });
}

// 创建启动脚本
function createLaunchScript(pythonCmd, hasFullDeps) {
    console.log('📝 创建启动脚本...');
    
    const launchScript = `#!/usr/bin/env node

/**
 * Markdown Hub 内置 Python 环境启动脚本
 * 自动使用打包的 Python 依赖
 */

const path = require('path');
const { spawn } = require('child_process');

const BUNDLE_DIR = path.join(__dirname, 'bundled');
const PYTHON_DEPS_DIR = path.join(BUNDLE_DIR, 'python-deps');
const BACKEND_DIR = path.join(BUNDLE_DIR, 'backend');

// 设置 Python 路径
process.env.PYTHONPATH = PYTHON_DEPS_DIR + path.delimiter + (process.env.PYTHONPATH || '');

// 启动 Python 脚本
function runPython(scriptPath, args = []) {
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn('${pythonCmd}', [scriptPath, ...args], {
            env: process.env,
            cwd: BACKEND_DIR,
            stdio: 'inherit'
        });
        
        pythonProcess.on('close', (code) => {
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(\`Python 进程退出，代码: \${code}\`));
            }
        });
        
        pythonProcess.on('error', (error) => {
            reject(error);
        });
    });
}

// 导出给 VS Code 扩展使用
module.exports = {
    runPython,
    BACKEND_DIR,
    PYTHON_DEPS_DIR,
    hasFullDeps: ${hasFullDeps}
};

// 如果直接运行此脚本
if (require.main === module) {
    const args = process.argv.slice(2);
    const scriptPath = path.join(BACKEND_DIR, 'cli.py');
    
    runPython(scriptPath, args)
        .then(() => process.exit(0))
        .catch((error) => {
            console.error('错误:', error.message);
            process.exit(1);
        });
}
`;

    fs.writeFileSync(path.join(__dirname, '..', 'bundled-launcher.js'), launchScript);
    console.log('✓ 启动脚本创建完成');
}

// 创建依赖信息文件
function createDepsInfo(hasFullDeps) {
    const depsInfo = {
        bundledAt: new Date().toISOString(),
        hasFullDeps: hasFullDeps,
        hasCairo: hasFullDeps,
        pythonDepsPath: './bundled/python-deps',
        backendPath: './bundled/backend',
        note: hasFullDeps ? 
            '包含完整依赖，支持所有功能包括 Cairo SVG 转换' : 
            '包含最小依赖，Cairo 功能将使用备用方案'
    };
    
    fs.writeFileSync(
        path.join(BUNDLE_DIR, 'deps-info.json'), 
        JSON.stringify(depsInfo, null, 2)
    );
    
    console.log('✓ 依赖信息文件创建完成');
}

// 主函数
async function main() {
    try {
        const pythonCmd = checkPython();
        const hasFullDeps = installDependencies(pythonCmd);
        copyBackendFiles();
        createLaunchScript(pythonCmd, hasFullDeps);
        createDepsInfo(hasFullDeps);
        
        console.log('🎉 打包完成！');
        console.log(`📦 打包目录: ${BUNDLE_DIR}`);
        console.log(`🐍 Python 依赖: ${PYTHON_DEPS_DIR}`);
        console.log(`🔧 功能状态: ${hasFullDeps ? '完整功能' : '基础功能'}`);
        
    } catch (error) {
        console.error('❌ 打包失败:', error.message);
        process.exit(1);
    }
}

// 如果直接运行此脚本
if (require.main === module) {
    main();
} 