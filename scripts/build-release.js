#!/usr/bin/env node

/**
 * 发布构建脚本
 * 自动化打包 Python 依赖和构建 VSCode 扩展
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('🚀 开始发布构建过程...');

// 检查必要的工具
function checkRequirements() {
    console.log('🔍 检查构建要求...');
    
    const requirements = [
        { cmd: 'node --version', name: 'Node.js' },
        { cmd: 'npm --version', name: 'npm' },
        { cmd: 'npx vsce --version', name: 'vsce (VS Code Extension Manager)' }
    ];
    
    for (const req of requirements) {
        try {
            execSync(req.cmd, { stdio: 'pipe' });
            console.log(`✓ ${req.name} 已安装`);
        } catch (error) {
            console.error(`❌ ${req.name} 未安装或不可用`);
            if (req.name === 'vsce (VS Code Extension Manager)') {
                console.log('请运行: npm install -g vsce');
            }
            process.exit(1);
        }
    }
}

// 清理旧的构建文件
function cleanup() {
    console.log('🧹 清理旧的构建文件...');
    
    const pathsToClean = [
        path.join(__dirname, '..', 'bundled'),
        path.join(__dirname, '..', 'bundled-launcher.js'),
        path.join(__dirname, '..', 'out')
    ];
    
    pathsToClean.forEach(cleanPath => {
        if (fs.existsSync(cleanPath)) {
            fs.rmSync(cleanPath, { recursive: true, force: true });
            console.log(`✓ 清理: ${path.basename(cleanPath)}`);
        }
    });
}

// 编译 TypeScript
function compileTypeScript() {
    console.log('🔨 编译 TypeScript...');
    
    try {
        execSync('npm run compile', { stdio: 'inherit' });
        console.log('✓ TypeScript 编译完成');
    } catch (error) {
        console.error('❌ TypeScript 编译失败');
        process.exit(1);
    }
}

// 打包 Python 依赖
function bundlePythonDeps() {
    console.log('📦 打包 Python 依赖...');
    
    try {
        execSync('npm run bundle-deps', { stdio: 'inherit' });
        console.log('✓ Python 依赖打包完成');
        
        // 检查打包结果
        const bundledPath = path.join(__dirname, '..', 'bundled');
        const depsInfoPath = path.join(bundledPath, 'deps-info.json');
        
        if (fs.existsSync(depsInfoPath)) {
            const depsInfo = JSON.parse(fs.readFileSync(depsInfoPath, 'utf8'));
            console.log(`📋 依赖信息:`);
            console.log(`   - 功能状态: ${depsInfo.hasFullDeps ? '完整功能' : '基础功能'}`);
            console.log(`   - Cairo 支持: ${depsInfo.hasCairo ? '是' : '否'}`);
            console.log(`   - 打包时间: ${depsInfo.bundledAt}`);
        }
        
    } catch (error) {
        console.error('❌ Python 依赖打包失败');
        process.exit(1);
    }
}

// 构建 VSCode 扩展
function buildExtension() {
    console.log('📦 构建 VSCode 扩展...');
    
    try {
        execSync('npx vsce package', { stdio: 'inherit' });
        console.log('✓ VSCode 扩展构建完成');
        
        // 查找生成的 .vsix 文件
        const files = fs.readdirSync(__dirname + '/..');
        const vsixFiles = files.filter(file => file.endsWith('.vsix'));
        
        if (vsixFiles.length > 0) {
            console.log(`📦 生成的扩展文件: ${vsixFiles[0]}`);
            
            // 显示文件大小
            const filePath = path.join(__dirname, '..', vsixFiles[0]);
            const stats = fs.statSync(filePath);
            const fileSizeInMB = (stats.size / (1024 * 1024)).toFixed(2);
            console.log(`📊 文件大小: ${fileSizeInMB} MB`);
        }
        
    } catch (error) {
        console.error('❌ VSCode 扩展构建失败');
        process.exit(1);
    }
}

// 生成构建报告
function generateBuildReport() {
    console.log('📋 生成构建报告...');
    
    const bundledPath = path.join(__dirname, '..', 'bundled');
    const depsInfoPath = path.join(bundledPath, 'deps-info.json');
    
    if (!fs.existsSync(depsInfoPath)) {
        console.log('⚠️ 未找到依赖信息文件');
        return;
    }
    
    const depsInfo = JSON.parse(fs.readFileSync(depsInfoPath, 'utf8'));
    const packageJson = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
    
    const buildReport = {
        extensionName: packageJson.name,
        extensionVersion: packageJson.version,
        buildTime: new Date().toISOString(),
        pythonDeps: {
            hasFullDeps: depsInfo.hasFullDeps,
            hasCairo: depsInfo.hasCairo,
            bundledAt: depsInfo.bundledAt,
            note: depsInfo.note
        },
        features: {
            markdownToDocx: true,
            markdownToPdf: true,
            markdownToHtml: true,
            markdownToPptx: true,
            officeToMarkdown: true,
            svgToPng: true,
            svgToPngHighQuality: depsInfo.hasCairo
        }
    };
    
    fs.writeFileSync(
        path.join(__dirname, '..', 'build-report.json'),
        JSON.stringify(buildReport, null, 2)
    );
    
    console.log('✓ 构建报告已生成: build-report.json');
    console.log('📊 功能支持情况:');
    console.log(`   - Markdown → DOCX: ✓`);
    console.log(`   - Markdown → PDF: ✓`);
    console.log(`   - Markdown → HTML: ✓`);
    console.log(`   - Markdown → PPTX: ✓`);
    console.log(`   - Office → Markdown: ✓`);
    console.log(`   - SVG → PNG: ✓`);
    console.log(`   - SVG → PNG (高质量): ${depsInfo.hasCairo ? '✓' : '❌ (将使用备用方案)'}`);
}

// 主函数
async function main() {
    try {
        checkRequirements();
        cleanup();
        compileTypeScript();
        bundlePythonDeps();
        buildExtension();
        generateBuildReport();
        
        console.log('\n🎉 发布构建完成！');
        console.log('📦 扩展已打包，包含所有 Python 依赖');
        console.log('🚀 用户安装后即可直接使用，无需额外配置');
        
    } catch (error) {
        console.error('❌ 构建失败:', error.message);
        process.exit(1);
    }
}

// 如果直接运行此脚本
if (require.main === module) {
    main();
} 