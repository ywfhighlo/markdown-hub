# 构建说明 - 打包 Python 依赖

本文档说明如何构建包含 Python 依赖的 VSCode 扩展，解决用户安装 Cairo 等依赖的困难。

## 🎯 目标

- 将 Python 依赖（包括 Cairo）打包到 VSCode 扩展中
- 用户安装扩展后无需额外配置，即可直接使用
- 支持自动回退机制，确保功能的可用性

## 📋 构建要求

在开始构建前，请确保您的系统已安装：

1. **Node.js** (v14 或更高版本)
2. **npm** 
3. **Python** (3.8 或更高版本)
4. **vsce** (VS Code Extension Manager)
   ```bash
   npm install -g vsce
   ```

## 🚀 快速构建

### 方法一：一键构建（推荐）
```bash
npm run build-release
```

这个命令会自动：
- 编译 TypeScript 代码
- 打包 Python 依赖
- 构建 VSCode 扩展
- 生成构建报告

### 方法二：分步构建
```bash
# 1. 编译 TypeScript
npm run compile

# 2. 打包 Python 依赖
npm run bundle-deps

# 3. 构建扩展
npm run package
```

## 📦 打包过程详解

### 1. Python 依赖打包
脚本会自动：
- 尝试安装完整依赖（包括 Cairo）
- 如果失败，回退到最小依赖
- 将依赖安装到 `bundled/python-deps/` 目录
- 复制后端文件到 `bundled/backend/` 目录

### 2. 自动回退机制
- **完整安装成功**：支持所有功能，包括高质量 SVG 转换
- **最小安装**：基础功能正常，SVG 转换使用备用方案

### 3. 扩展打包
- TypeScript 代码会检测是否有打包的依赖
- 如果有，优先使用打包的依赖
- 如果没有，回退到用户环境

## 🔧 配置选项

### 环境变量
- `PYTHONPATH`: 会自动设置为打包的依赖路径
- `NODE_ENV`: 可设置为 `production` 进行优化

### 自定义打包
如果需要自定义打包过程，可以修改 `scripts/bundle-deps.js`：

```javascript
// 修改要打包的文件
const filesToCopy = [
    'cli.py',
    'converters',
    'your-custom-files'  // 添加您的文件
];
```

## 📊 构建结果

构建完成后，您会看到：

1. **扩展文件**: `markdown-hub-x.x.x.vsix`
2. **构建报告**: `build-report.json`
3. **打包目录**: `bundled/`

### 构建报告示例
```json
{
  "extensionName": "markdown-hub",
  "extensionVersion": "0.2.0",
  "buildTime": "2024-01-01T00:00:00.000Z",
  "pythonDeps": {
    "hasFullDeps": true,
    "hasCairo": true,
    "bundledAt": "2024-01-01T00:00:00.000Z",
    "note": "包含完整依赖，支持所有功能包括 Cairo SVG 转换"
  },
  "features": {
    "markdownToDocx": true,
    "markdownToPdf": true,
    "markdownToHtml": true,
    "markdownToPptx": true,
    "officeToMarkdown": true,
    "svgToPng": true,
    "svgToPngHighQuality": true
  }
}
```

## 🐛 常见问题

### Q1: Python 依赖打包失败
**解决方案**：
1. 检查 Python 版本（需要 3.8+）
2. 确保网络连接正常
3. 如果 Cairo 安装失败，脚本会自动回退到最小依赖

### Q2: 扩展文件过大
**解决方案**：
1. 使用最小依赖构建：修改 `scripts/bundle-deps.js` 强制使用 `requirements-minimal.txt`
2. 清理不必要的文件

### Q3: 用户环境中 Python 路径问题
**解决方案**：
扩展会自动检测并使用打包的依赖，无需用户配置 Python 路径。

## 🚀 发布流程

1. **构建扩展**：
   ```bash
   npm run build-release
   ```

2. **测试扩展**：
   ```bash
   code --install-extension markdown-hub-x.x.x.vsix
   ```

3. **发布到市场**：
   ```bash
   vsce publish
   ```

## 📈 优势

- **零配置**：用户安装后即可使用
- **自动回退**：确保功能可用性
- **跨平台**：支持 Windows、macOS、Linux
- **完整功能**：包含所有 Python 依赖

## 🔄 更新依赖

如果需要更新 Python 依赖：

1. 修改 `backend/requirements.txt`
2. 重新运行构建：`npm run build-release`
3. 测试新版本的功能

## 📞 技术支持

如果在构建过程中遇到问题：
1. 查看构建日志
2. 检查 `build-report.json` 文件
3. 提交 GitHub Issue 并附上错误信息 