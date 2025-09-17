# Markdown Hub 后端Java迁移设计方案

## 1. 架构设计

### 1.1 整体架构

```
┌─────────────────────────────────────────────────────────────┐
│                    VSCode Extension (TypeScript)            │
│  ┌─────────────────┐  ┌─────────────────┐  ┌──────────────┐ │
│  │  commandHandler │  │   extension.ts  │  │ pythonService│ │
│  └─────────────────┘  └─────────────────┘  └──────────────┘ │
└─────────────────────────┬───────────────────────────────────┘
                          │ JSON-RPC / Process Communication
                          ▼
┌─────────────────────────────────────────────────────────────┐
│                Java Backend (markdown-hub.jar)             │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                    CLI Entry Point                     │ │
│  │  ┌─────────────────┐  ┌─────────────────┐             │ │
│  │  │  ArgumentParser │  │  CommandRouter  │             │ │
│  │  └─────────────────┘  └─────────────────┘             │ │
│  └─────────────────────────────────────────────────────────┘ │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                 Converter Factory                      │ │
│  │  ┌─────────────────┐  ┌─────────────────┐             │ │
│  │  │ ConverterRegistry│  │  BaseConverter  │             │ │
│  │  └─────────────────┘  └─────────────────┘             │ │
│  └─────────────────────────────────────────────────────────┘ │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                    Converters                          │ │
│  │  ┌─────────────────┐  ┌─────────────────┐             │ │
│  │  │DiagramConverter │  │ OfficeConverter │             │ │
│  │  │  - PlantUML     │  │  - MD→Office    │             │ │
│  │  │  - SVG (Batik)  │  │  - Office→MD    │             │ │
│  │  │  - Mermaid      │  │  - PDF Processing│             │ │
│  │  └─────────────────┘  └─────────────────┘             │ │
│  └─────────────────────────────────────────────────────────┘ │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │                 External Tools                         │ │
│  │  ┌─────────────────┐  ┌─────────────────┐             │ │
│  │  │   Pandoc        │  │   System Tools  │             │ │
│  │  │   Integration   │  │   (rsvg, etc.)  │             │ │
│  │  └─────────────────┘  └─────────────────┘             │ │
│  └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

### 1.2 模块设计

#### 1.2.1 核心模块
- **CLI模块**：命令行参数解析和路由
- **转换器工厂**：转换器的注册和实例化
- **基础转换器**：所有转换器的抽象基类
- **具体转换器**：各种格式转换的具体实现
- **工具集成模块**：外部工具的调用和管理
- **配置管理**：转换参数和系统配置
- **日志系统**：统一的日志记录

#### 1.2.2 包结构
```
com.markdownhub
├── cli/
│   ├── CliApplication.java
│   ├── ArgumentParser.java
│   └── CommandRouter.java
├── converter/
│   ├── base/
│   │   ├── BaseConverter.java
│   │   ├── ConverterFactory.java
│   │   └── ConverterRegistry.java
│   ├── diagram/
│   │   ├── DiagramConverter.java
│   │   ├── PlantUMLConverter.java
│   │   ├── SvgConverter.java
│   │   └── MermaidConverter.java
│   └── office/
│       ├── MdToOfficeConverter.java
│       ├── OfficeToMdConverter.java
│       └── PdfConverter.java
├── external/
│   ├── PandocIntegration.java
│   ├── SystemToolManager.java
│   └── ProcessExecutor.java
├── config/
│   ├── Configuration.java
│   └── ConverterConfig.java
└── util/
    ├── FileUtils.java
    ├── LoggerFactory.java
    └── DependencyChecker.java
```

## 2. 技术选型

### 2.1 核心技术栈

#### 2.1.1 Java平台
- **Java版本**：Java 11 LTS（平衡兼容性和功能）
- **构建工具**：Maven 3.8+
- **打包方式**：Maven Shade Plugin（Fat JAR）

#### 2.1.2 核心依赖库
```xml
<dependencies>
    <!-- 命令行参数解析 -->
    <dependency>
        <groupId>info.picocli</groupId>
        <artifactId>picocli</artifactId>
        <version>4.7.5</version>
    </dependency>
    
    <!-- Office文档处理 -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.4</version>
    </dependency>
    
    <!-- Markdown处理 -->
    <dependency>
        <groupId>com.vladsch.flexmark</groupId>
        <artifactId>flexmark-all</artifactId>
        <version>0.64.8</version>
    </dependency>
    
    <!-- PDF处理 -->
    <dependency>
        <groupId>org.apache.pdfbox</groupId>
        <artifactId>pdfbox</artifactId>
        <version>2.0.29</version>
    </dependency>
    
    <!-- JSON处理 -->
    <dependency>
        <groupId>com.fasterxml.jackson.core</groupId>
        <artifactId>jackson-databind</artifactId>
        <version>2.15.2</version>
    </dependency>
    
    <!-- 日志框架 -->
    <dependency>
        <groupId>ch.qos.logback</groupId>
        <artifactId>logback-classic</artifactId>
        <version>1.4.11</version>
    </dependency>
    
    <!-- 测试框架 -->
    <dependency>
        <groupId>org.junit.jupiter</groupId>
        <artifactId>junit-jupiter</artifactId>
        <version>5.10.0</version>
        <scope>test</scope>
    </dependency>
</dependencies>
```

### 2.2 现有工具集成

#### 2.2.1 保留的JAR工具
- **PlantUML**：`tools/plantuml.jar`
- **Batik**：`tools/batik-lib/` 目录下的所有JAR

#### 2.2.2 外部工具依赖
- **Pandoc**：通过ProcessBuilder调用
- **mmdc**：Mermaid转换（如果找不到Java替代）

## 3. 详细设计

### 3.1 CLI接口设计

#### 3.1.1 命令行参数
保持与Python版本完全一致的参数格式：

```bash
# 图表转PNG
java -jar markdown-hub.jar diagram-to-png input.svg output.png

# Markdown转Office
java -jar markdown-hub.jar md-to-office input.md output.docx --format docx

# Office转Markdown
java -jar markdown-hub.jar office-to-md input.docx output.md

# 批量转换
java -jar markdown-hub.jar diagram-to-png /path/to/diagrams/ /path/to/output/
```

#### 3.1.2 参数解析
使用PicoCLI框架实现类型安全的参数解析：

```java
@Command(name = "markdown-hub", mixinStandardHelpOptions = true)
public class CliApplication implements Callable<Integer> {
    
    @Parameters(index = "0", description = "转换类型")
    private String converterType;
    
    @Parameters(index = "1", description = "输入文件或目录")
    private String inputPath;
    
    @Parameters(index = "2", description = "输出文件或目录")
    private String outputPath;
    
    @Option(names = {"--format", "-f"}, description = "输出格式")
    private String format;
    
    @Option(names = {"--config", "-c"}, description = "配置文件路径")
    private String configPath;
}
```

### 3.2 转换器架构

#### 3.2.1 基础转换器接口
```java
public abstract class BaseConverter {
    protected final Configuration config;
    protected final Logger logger;
    
    public BaseConverter(Configuration config) {
        this.config = config;
        this.logger = LoggerFactory.getLogger(this.getClass());
    }
    
    public abstract List<String> convert(String inputPath) throws ConversionException;
    
    protected abstract boolean isValidInput(String inputPath, List<String> expectedExtensions);
    
    protected abstract void checkDependencies() throws DependencyException;
}
```

#### 3.2.2 转换器注册机制
```java
@Component
public class ConverterRegistry {
    private final Map<String, Class<? extends BaseConverter>> converters = new HashMap<>();
    
    public void register(String type, Class<? extends BaseConverter> converterClass) {
        converters.put(type, converterClass);
    }
    
    public BaseConverter create(String type, Configuration config) {
        Class<? extends BaseConverter> clazz = converters.get(type);
        if (clazz == null) {
            throw new IllegalArgumentException("Unknown converter type: " + type);
        }
        return instantiate(clazz, config);
    }
}
```

### 3.3 具体转换器实现

#### 3.3.1 PlantUML转换器
```java
public class PlantUMLConverter extends BaseConverter {
    private static final String PLANTUML_JAR = "tools/plantuml.jar";
    
    @Override
    public List<String> convert(String inputPath) throws ConversionException {
        checkDependencies();
        
        List<String> command = Arrays.asList(
            "java", "-jar", PLANTUML_JAR,
            "-tpng", inputPath
        );
        
        ProcessResult result = ProcessExecutor.execute(command);
        if (result.getExitCode() != 0) {
            throw new ConversionException("PlantUML conversion failed: " + result.getError());
        }
        
        return findGeneratedFiles(inputPath);
    }
    
    @Override
    protected void checkDependencies() throws DependencyException {
        if (!Files.exists(Paths.get(PLANTUML_JAR))) {
            throw new DependencyException("PlantUML JAR not found: " + PLANTUML_JAR);
        }
        
        if (!SystemToolManager.isJavaAvailable()) {
            throw new DependencyException("Java runtime not available");
        }
    }
}
```

#### 3.3.2 SVG转换器（Batik）
```java
public class SvgConverter extends BaseConverter {
    private static final String BATIK_LIB_DIR = "tools/batik-lib";
    
    @Override
    public List<String> convert(String inputPath) throws ConversionException {
        checkDependencies();
        
        String classpath = buildBatikClasspath();
        List<String> command = Arrays.asList(
            "java", "-cp", classpath,
            "org.apache.batik.apps.rasterizer.Main",
            "-d", getOutputPath(inputPath),
            inputPath
        );
        
        ProcessResult result = ProcessExecutor.execute(command);
        if (result.getExitCode() != 0) {
            throw new ConversionException("Batik conversion failed: " + result.getError());
        }
        
        return Arrays.asList(getOutputPath(inputPath));
    }
    
    private String buildBatikClasspath() {
        Path libDir = Paths.get(BATIK_LIB_DIR);
        return Files.list(libDir)
            .filter(path -> path.toString().endsWith(".jar"))
            .map(Path::toString)
            .collect(Collectors.joining(System.getProperty("path.separator")));
    }
}
```

#### 3.3.3 Markdown转Office转换器
```java
public class MdToOfficeConverter extends BaseConverter {
    
    @Override
    public List<String> convert(String inputPath) throws ConversionException {
        String format = config.getOutputFormat();
        
        switch (format.toLowerCase()) {
            case "docx":
                return convertToDocx(inputPath);
            case "pptx":
                return convertToPptx(inputPath);
            case "pdf":
                return convertToPdf(inputPath);
            case "html":
                return convertToHtml(inputPath);
            default:
                throw new ConversionException("Unsupported format: " + format);
        }
    }
    
    private List<String> convertToDocx(String inputPath) throws ConversionException {
        // 使用Pandoc进行转换
        List<String> command = Arrays.asList(
            "pandoc",
            "-f", "markdown",
            "-t", "docx",
            "-o", getOutputPath(inputPath, ".docx"),
            inputPath
        );
        
        ProcessResult result = ProcessExecutor.execute(command);
        if (result.getExitCode() != 0) {
            throw new ConversionException("Pandoc conversion failed: " + result.getError());
        }
        
        return Arrays.asList(getOutputPath(inputPath, ".docx"));
    }
    
    private List<String> convertToPptx(String inputPath) throws ConversionException {
        // 使用Apache POI直接生成PPTX
        try {
            String content = Files.readString(Paths.get(inputPath));
            Document document = FlexmarkHtmlConverter.builder().build().convert(content);
            
            XMLSlideShow ppt = new XMLSlideShow();
            // 处理Markdown内容，生成幻灯片
            processMarkdownToPptx(document, ppt);
            
            String outputPath = getOutputPath(inputPath, ".pptx");
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                ppt.write(out);
            }
            
            return Arrays.asList(outputPath);
        } catch (IOException e) {
            throw new ConversionException("PPTX conversion failed", e);
        }
    }
}
```

### 3.4 外部工具集成

#### 3.4.1 进程执行器
```java
public class ProcessExecutor {
    private static final Logger logger = LoggerFactory.getLogger(ProcessExecutor.class);
    
    public static ProcessResult execute(List<String> command) throws ConversionException {
        return execute(command, null, 60); // 默认60秒超时
    }
    
    public static ProcessResult execute(List<String> command, Path workingDir, int timeoutSeconds) 
            throws ConversionException {
        try {
            ProcessBuilder pb = new ProcessBuilder(command);
            if (workingDir != null) {
                pb.directory(workingDir.toFile());
            }
            
            pb.redirectErrorStream(true);
            Process process = pb.start();
            
            boolean finished = process.waitFor(timeoutSeconds, TimeUnit.SECONDS);
            if (!finished) {
                process.destroyForcibly();
                throw new ConversionException("Process timeout: " + String.join(" ", command));
            }
            
            String output = new String(process.getInputStream().readAllBytes());
            return new ProcessResult(process.exitValue(), output);
            
        } catch (IOException | InterruptedException e) {
            throw new ConversionException("Process execution failed", e);
        }
    }
}
```

#### 3.4.2 系统工具管理
```java
public class SystemToolManager {
    private static final Map<String, Boolean> toolCache = new ConcurrentHashMap<>();
    
    public static boolean isToolAvailable(String toolName) {
        return toolCache.computeIfAbsent(toolName, SystemToolManager::checkTool);
    }
    
    private static boolean checkTool(String toolName) {
        try {
            ProcessResult result = ProcessExecutor.execute(
                Arrays.asList(toolName, "--version"), null, 5
            );
            return result.getExitCode() == 0;
        } catch (ConversionException e) {
            return false;
        }
    }
    
    public static boolean isJavaAvailable() {
        return isToolAvailable("java");
    }
    
    public static boolean isPandocAvailable() {
        return isToolAvailable("pandoc");
    }
}
```

## 4. 实施计划

### 4.1 第一阶段：基础架构（2-3周）

#### 里程碑1.1：项目搭建
- [ ] 创建Maven项目结构
- [ ] 配置构建脚本和依赖
- [ ] 设置CI/CD流水线
- [ ] 建立代码规范和质量检查

#### 里程碑1.2：核心框架
- [ ] 实现CLI参数解析
- [ ] 实现转换器工厂和注册机制
- [ ] 实现基础转换器抽象类
- [ ] 实现配置管理和日志系统

#### 里程碑1.3：外部工具集成
- [ ] 实现进程执行器
- [ ] 实现系统工具检查
- [ ] 实现依赖检查机制

### 4.2 第二阶段：图表转换（2-3周）

#### 里程碑2.1：PlantUML转换
- [ ] 实现PlantUML转换器
- [ ] 集成现有plantuml.jar
- [ ] 添加单元测试和集成测试

#### 里程碑2.2：SVG转换
- [ ] 实现Batik SVG转换器
- [ ] 集成现有batik-lib
- [ ] 添加rsvg-convert备选方案

#### 里程碑2.3：其他图表格式
- [ ] 研究Mermaid Java替代方案
- [ ] 实现Draw.io转换（如果可行）
- [ ] 完善图表转换测试

### 4.3 第三阶段：Office转换（3-4周）

#### 里程碑3.1：Markdown转Office
- [ ] 实现Markdown转DOCX
- [ ] 实现Markdown转PPTX
- [ ] 实现Markdown转HTML
- [ ] 集成Pandoc调用

#### 里程碑3.2：Office转Markdown
- [ ] 实现DOCX转Markdown
- [ ] 实现PPTX转Markdown
- [ ] 实现PDF转Markdown
- [ ] 实现Excel转Markdown

#### 里程碑3.3：PDF处理
- [ ] 实现PDF文本提取
- [ ] 实现OCR功能集成
- [ ] 优化PDF转换质量

### 4.4 第四阶段：集成测试（1-2周）

#### 里程碑4.1：功能测试
- [ ] 完整功能回归测试
- [ ] 性能基准测试
- [ ] 跨平台兼容性测试

#### 里程碑4.2：VSCode集成
- [ ] 验证CLI接口兼容性
- [ ] 测试VSCode扩展集成
- [ ] 优化错误处理和用户体验

### 4.5 第五阶段：部署和文档（1周）

#### 里程碑5.1：打包部署
- [ ] 配置Fat JAR打包
- [ ] 优化JAR大小
- [ ] 创建安装脚本

#### 里程碑5.2：文档和发布
- [ ] 更新用户文档
- [ ] 创建迁移指南
- [ ] 发布Java版本

## 5. 质量保证

### 5.1 测试策略

#### 5.1.1 单元测试
- 每个转换器的核心逻辑测试
- 工具类和辅助方法测试
- 异常处理和边界条件测试
- 目标覆盖率：>80%

#### 5.1.2 集成测试
- 端到端转换流程测试
- 外部工具集成测试
- 文件I/O和格式验证测试
- 跨平台兼容性测试

#### 5.1.3 性能测试
- 转换速度基准测试
- 内存使用监控
- 大文件处理测试
- 并发转换测试

### 5.2 代码质量

#### 5.2.1 静态分析
- SpotBugs：查找潜在bug
- PMD：代码质量检查
- Checkstyle：代码风格检查
- SonarQube：综合质量分析

#### 5.2.2 代码审查
- 所有代码变更需要审查
- 关注设计模式和最佳实践
- 确保异常处理和资源管理
- 验证测试覆盖率

## 6. 风险缓解

### 6.1 技术风险缓解

#### 6.1.1 依赖库风险
- **问题**：某些Python库没有Java等价物
- **缓解**：提前调研和验证关键库，准备备选方案
- **应急**：保留外部工具调用作为后备

#### 6.1.2 性能风险
- **问题**：Java版本性能可能不如Python
- **缓解**：早期性能测试，优化关键路径
- **应急**：使用JVM调优和并行处理

#### 6.1.3 兼容性风险
- **问题**：输出格式可能与Python版本不一致
- **缓解**：详细的对比测试，确保输出一致性
- **应急**：提供兼容性配置选项

### 6.2 项目风险缓解

#### 6.2.1 进度风险
- **问题**：开发时间可能超出预期
- **缓解**：分阶段实施，优先核心功能
- **应急**：调整功能范围，延后非关键特性

#### 6.2.2 质量风险
- **问题**：可能出现功能回归
- **缓解**：充分的测试覆盖，自动化测试
- **应急**：快速修复机制，回滚计划

## 7. 成功标准

### 7.1 功能完整性
- [ ] 所有现有转换功能正常工作
- [ ] CLI接口100%兼容
- [ ] 输出质量与Python版本一致
- [ ] VSCode扩展无需修改

### 7.2 性能指标
- [ ] 转换速度不低于Python版本
- [ ] 启动时间<3秒
- [ ] 内存使用<512MB（正常文件）
- [ ] 支持>100MB的大文件

### 7.3 部署简化
- [ ] 单JAR文件部署
- [ ] 依赖减少到<5个外部工具
- [ ] 安装成功率>95%
- [ ] 跨平台一致性100%

### 7.4 维护性
- [ ] 代码测试覆盖率>80%
- [ ] 文档完整性>90%
- [ ] 代码质量评分>A级
- [ ] 新功能开发效率提升>25%