# Markdown Hub Java Backend

一个功能强大的文档转换工具，支持多种格式之间的转换，包括Markdown、Office文档和图表。

## 功能特性

- **图表转换**: 支持PlantUML和Mermaid图表转换为PNG格式
- **Markdown转Office**: 将Markdown文档转换为Word (.docx) 和PowerPoint (.pptx) 格式
- **Office转Markdown**: 将Word和PowerPoint文档转换为Markdown格式
- **命令行界面**: 基于PicoCLI的友好命令行界面
- **配置管理**: 灵活的配置系统，支持JSON配置文件
- **日志系统**: 完整的日志记录和性能监控
- **依赖检查**: 自动检查外部工具依赖

## 系统要求

- Java 11 或更高版本
- Maven 3.6 或更高版本

### 外部依赖工具

- **PlantUML**: 用于UML图表转换
- **Batik**: 用于SVG到PNG转换
- **Mermaid CLI**: 用于Mermaid图表转换 (可选)
- **Pandoc**: 用于高级文档转换 (可选)

## 构建和安装

### 1. 克隆项目

```bash
git clone <repository-url>
cd markdown-hub/java-backend
```

### 2. 构建项目

```bash
mvn clean compile
```

### 3. 运行测试

```bash
mvn test
```

### 4. 打包应用

```bash
mvn package
```

这将在 `target/` 目录下生成可执行的JAR文件。

## 使用方法

### 基本语法

```bash
java -jar target/markdown-hub-1.0.0.jar [OPTIONS] <input-file> <output-file>
```

### 命令行选项

- `-t, --type <TYPE>`: 指定转换类型 (diagram-to-png, md-to-office, office-to-md)
- `-c, --config <FILE>`: 指定配置文件路径
- `-v, --verbose`: 启用详细日志输出
- `-q, --quiet`: 静默模式，只输出错误信息
- `--check-deps`: 检查外部依赖工具
- `-h, --help`: 显示帮助信息

### 使用示例

#### 1. 图表转PNG

```bash
# PlantUML图表转PNG
java -jar target/markdown-hub-1.0.0.jar -t diagram-to-png diagram.puml diagram.png

# Mermaid图表转PNG
java -jar target/markdown-hub-1.0.0.jar -t diagram-to-png flowchart.mmd flowchart.png
```

#### 2. Markdown转Office

```bash
# Markdown转Word
java -jar target/markdown-hub-1.0.0.jar -t md-to-office document.md document.docx

# Markdown转PowerPoint
java -jar target/markdown-hub-1.0.0.jar -t md-to-office presentation.md presentation.pptx
```

#### 3. Office转Markdown

```bash
# Word转Markdown
java -jar target/markdown-hub-1.0.0.jar -t office-to-md document.docx document.md

# PowerPoint转Markdown
java -jar target/markdown-hub-1.0.0.jar -t office-to-md presentation.pptx presentation.md
```

#### 4. 检查依赖

```bash
java -jar target/markdown-hub-1.0.0.jar --check-deps
```

## 配置文件

应用支持JSON格式的配置文件，可以自定义各种转换参数：

```json
{
  "outputDirectory": "./output",
  "tempDirectory": "./temp",
  "logLevel": "INFO",
  "maxFileSize": 104857600,
  "converters": {
    "diagram": {
      "plantUmlJarPath": "/path/to/plantuml.jar",
      "batikJarPath": "/path/to/batik-rasterizer.jar",
      "mermaidCliPath": "mmdc",
      "outputFormat": "png",
      "dpi": 300
    },
    "office": {
      "defaultFont": "Arial",
      "fontSize": 12,
      "pageMargin": 72
    }
  },
  "performance": {
    "enableMetrics": true,
    "timeoutSeconds": 300,
    "maxConcurrentJobs": 4
  }
}
```

## 项目结构

```
src/main/java/com/markdownhub/
├── CliApplication.java              # 主应用入口
├── config/
│   ├── Configuration.java           # 配置管理
│   └── LoggingConfig.java          # 日志配置
├── converter/
│   ├── base/
│   │   ├── BaseConverter.java      # 转换器基类
│   │   ├── ConversionException.java # 转换异常
│   │   └── DependencyException.java # 依赖异常
│   ├── diagram/
│   │   └── DiagramToPngConverter.java # 图表转PNG转换器
│   ├── markdown/
│   │   └── MarkdownToOfficeConverter.java # Markdown转Office转换器
│   ├── office/
│   │   └── OfficeToMarkdownConverter.java # Office转Markdown转换器
│   └── ConverterFactory.java       # 转换器工厂
└── util/
    ├── ProcessExecutor.java         # 进程执行器
    └── SystemToolChecker.java       # 系统工具检查
```

## 开发指南

### 添加新的转换器

1. 继承 `BaseConverter` 类
2. 实现必要的抽象方法
3. 在 `ConverterFactory` 中注册新转换器
4. 添加相应的测试用例

### 代码质量

项目使用以下工具确保代码质量：

- **SpotBugs**: 静态代码分析
- **Checkstyle**: 代码风格检查
- **JaCoCo**: 代码覆盖率报告

运行质量检查：

```bash
mvn spotbugs:check
mvn checkstyle:check
mvn jacoco:report
```

## 故障排除

### 常见问题

1. **依赖工具未找到**
   - 运行 `--check-deps` 检查依赖状态
   - 确保外部工具已正确安装并在PATH中

2. **内存不足错误**
   - 增加JVM堆内存：`java -Xmx2g -jar ...`
   - 检查输入文件大小是否超过限制

3. **转换失败**
   - 启用详细日志：`-v` 选项
   - 检查输入文件格式是否正确
   - 查看日志文件获取详细错误信息

### 日志文件

应用日志默认输出到：
- 控制台输出
- `logs/markdown-hub.log` (如果配置了文件输出)

## 许可证

本项目采用 MIT 许可证。详见 LICENSE 文件。

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。

## 更新日志

### v1.0.0
- 初始版本发布
- 支持图表转PNG转换
- 支持Markdown与Office文档互转
- 完整的CLI界面和配置系统