#!/usr/bin/env node

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
        const pythonProcess = spawn('python', [scriptPath, ...args], {
            env: process.env,
            cwd: BACKEND_DIR,
            stdio: 'inherit'
        });
        
        pythonProcess.on('close', (code) => {
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(`Python 进程退出，代码: ${code}`));
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
    hasFullDeps: true
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
