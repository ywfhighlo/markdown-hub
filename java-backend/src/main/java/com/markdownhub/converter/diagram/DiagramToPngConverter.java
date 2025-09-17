package com.markdownhub.converter.diagram;

import com.markdownhub.config.Configuration;
import com.markdownhub.converter.base.BaseConverter;
import com.markdownhub.converter.base.ConversionException;
import com.markdownhub.converter.base.DependencyException;
import com.markdownhub.external.ProcessExecutor;
import com.markdownhub.external.SystemToolChecker;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.TranscoderException;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 图表转PNG转换器
 * 支持PlantUML和Mermaid图表转换为PNG图片
 */
public class DiagramToPngConverter extends BaseConverter {
    
    private static final Logger logger = LoggerFactory.getLogger(DiagramToPngConverter.class);
    private static final Logger perfLogger = LoggerFactory.getLogger("com.markdownhub.performance.DiagramToPngConverter");
    
    // PlantUML图表类型检测模式
    private static final Pattern PLANTUML_PATTERN = Pattern.compile(
            "@start(uml|sequence|usecase|class|activity|component|state|object|deployment|timing|network|salt|mindmap|wbs|gantt|json|yaml|creole|flow|board|git)", 
            Pattern.CASE_INSENSITIVE);
    
    // Mermaid图表类型检测模式
    private static final Pattern MERMAID_PATTERN = Pattern.compile(
            "^\\s*(graph|sequenceDiagram|classDiagram|stateDiagram|erDiagram|journey|gantt|pie|flowchart|gitgraph|mindmap|timeline|quadrantChart|xyChart|block|packet)", 
            Pattern.CASE_INSENSITIVE | Pattern.MULTILINE);
    
    private final ProcessExecutor processExecutor;
    private final SystemToolChecker toolChecker;
    
    public DiagramToPngConverter(Configuration config) {
        super(config);
        this.processExecutor = new ProcessExecutor(config);
        this.toolChecker = new SystemToolChecker(config);
    }
    
    @Override
    public String getConverterType() {
        return "diagram-to-png";
    }
    
    @Override
    public List<String> getSupportedInputFormats() {
        return Arrays.asList("puml", "plantuml", "mmd", "mermaid", "txt");
    }
    
    @Override
    public List<String> getSupportedOutputFormats() {
        return Arrays.asList("png");
    }
    
    @Override
    protected void checkDependencies() throws DependencyException {
        // 检查Java环境（必需）
        if (!toolChecker.isToolAvailable("java")) {
            throw new DependencyException("Java环境不可用，无法执行图表转换");
        }
        
        // 检查PlantUML JAR（用于PlantUML图表）
        boolean plantUmlAvailable = toolChecker.isToolAvailable("plantuml");
        
        // 检查Batik库（用于SVG转PNG）
        boolean batikAvailable = toolChecker.isToolAvailable("batik");
        
        // 检查Mermaid CLI（用于Mermaid图表）
        boolean mermaidAvailable = toolChecker.isToolAvailable("mmdc");
        
        if (!plantUmlAvailable && !mermaidAvailable) {
            throw new DependencyException("PlantUML JAR和Mermaid CLI都不可用，无法进行图表转换");
        }
        
        if (!batikAvailable) {
            logger.warn("Batik库不可用，将使用PlantUML内置PNG输出（质量可能较低）");
        }
        
        logger.debug("依赖检查完成 - PlantUML: {}, Batik: {}, Mermaid: {}", 
                plantUmlAvailable, batikAvailable, mermaidAvailable);
    }
    
    @Override
    protected void performConversion(Path inputPath, Path outputPath) throws ConversionException {
        try {
            long startTime = System.nanoTime();
            
            // 读取输入文件内容
            String content = Files.readString(inputPath, StandardCharsets.UTF_8);
            
            // 检测图表类型并转换
            DiagramType diagramType = detectDiagramType(inputPath);
            logger.debug("检测到图表类型: {}", diagramType);
            convertDiagram(inputPath, outputPath, diagramType);
            
            if (perfLogger.isDebugEnabled()) {
                long duration = System.nanoTime() - startTime;
                perfLogger.debug("图表转换完成: {} -> {} (耗时: {:.2f}ms)", 
                        inputPath.getFileName(), outputPath.getFileName(), duration / 1_000_000.0);
            }
            
        } catch (IOException e) {
            throw new ConversionException("读取输入文件失败: " + inputPath, e);
        }
    }
    
    /**
     * 检测图表类型
     * 
     * @param inputPath 输入文件路径
     * @return 图表类型
     * @throws IOException 读取文件失败
     */
    private DiagramType detectDiagramType(Path inputPath) throws IOException {
        String content = Files.readString(inputPath, StandardCharsets.UTF_8);
        // 检查PlantUML标记
        Matcher plantUmlMatcher = PLANTUML_PATTERN.matcher(content);
        if (plantUmlMatcher.find()) {
            return DiagramType.PLANTUML;
        }
        
        // 检查Mermaid标记
        Matcher mermaidMatcher = MERMAID_PATTERN.matcher(content);
        if (mermaidMatcher.find()) {
            return DiagramType.MERMAID;
        }
        
        return DiagramType.UNKNOWN;
    }
    
    /**
     * 转换图表
     * 
     * @param inputPath 输入路径
     * @param outputPath 输出路径
     * @param diagramType 图表类型
     * @throws ConversionException 转换失败
     * @throws IOException IO异常
     */
    private void convertDiagram(Path inputPath, Path outputPath, DiagramType diagramType) throws ConversionException, IOException {
        String content = Files.readString(inputPath, StandardCharsets.UTF_8);
        
        switch (diagramType) {
            case PLANTUML:
                convertPlantUmlDiagram(content, outputPath);
                break;
            case MERMAID:
                convertMermaidDiagram(content, outputPath);
                break;
            case UNKNOWN:
                logger.warn("无法确定图表类型，尝试作为PlantUML处理: {}", inputPath);
                convertPlantUmlDiagram(content, outputPath);
                break;
        }
    }
    
    /**
     * 转换PlantUML图表
     * 
     * @param content 图表内容
     * @param outputPath 输出路径
     * @throws ConversionException 转换失败
     */
    private void convertPlantUmlDiagram(String content, Path outputPath) throws ConversionException {
        if (!toolChecker.isToolAvailable("plantuml")) {
            throw new ConversionException("PlantUML不可用，无法转换PlantUML图表");
        }
        
        try {
            // 检查是否使用Batik进行高质量转换
            boolean useBatik = toolChecker.isToolAvailable("batik") && 
                    config.getConverterConfig("use_batik_for_plantuml", true);
            
            if (useBatik) {
                convertPlantUmlWithBatik(content, outputPath);
            } else {
                convertPlantUmlDirect(content, outputPath);
            }
            
        } catch (IOException e) {
            throw new ConversionException("PlantUML转换失败", e);
        }
    }
    
    /**
     * 使用Batik进行高质量PlantUML转换
     * 
     * @param content 图表内容
     * @param outputPath 输出路径
     * @throws IOException 转换失败
     * @throws ConversionException 转换失败
     */
    private void convertPlantUmlWithBatik(String content, Path outputPath) throws IOException, ConversionException {
        // 创建临时SVG文件
        Path tempSvgPath = Files.createTempFile("plantuml_", ".svg");
        
        try {
            // 使用PlantUML生成SVG
            List<String> plantUmlArgs = Arrays.asList(
                    "-tsvg",
                    "-pipe",
                    "-charset", "UTF-8"
            );
            
            ProcessExecutor.ExecutionResult result = processExecutor.executeJava(
                    config.getPlantUmlJarPath(), plantUmlArgs);
            
            if (!result.isSuccess()) {
                throw new ConversionException("PlantUML SVG生成失败: " + result.getStderr());
            }
            
            // 将SVG内容写入临时文件
            Files.writeString(tempSvgPath, result.getStdout(), StandardCharsets.UTF_8);
            
            // 使用Batik将SVG转换为PNG
            convertSvgToPngWithBatik(tempSvgPath, outputPath);
            
        } finally {
            // 清理临时文件
            try {
                Files.deleteIfExists(tempSvgPath);
            } catch (IOException e) {
                logger.warn("清理临时SVG文件失败: {}", tempSvgPath, e);
            }
        }
    }
    
    /**
     * 直接使用PlantUML生成PNG
     * 
     * @param content 图表内容
     * @param outputPath 输出路径
     * @throws IOException 转换失败
     * @throws ConversionException 转换失败
     */
    private void convertPlantUmlDirect(String content, Path outputPath) throws IOException, ConversionException {
        List<String> plantUmlArgs = Arrays.asList(
                "-tpng",
                "-pipe",
                "-charset", "UTF-8"
        );
        
        ProcessExecutor.ExecutionResult result = processExecutor.executeJava(
                config.getPlantUmlJarPath(), plantUmlArgs);
        
        if (!result.isSuccess()) {
            throw new ConversionException("PlantUML PNG生成失败: " + result.getStderr());
        }
        
        // 将PNG数据写入输出文件
        byte[] pngData = result.getStdout().getBytes(StandardCharsets.ISO_8859_1);
        Files.write(outputPath, pngData);
    }
    
    /**
     * 转换Mermaid图表
     * 
     * @param content 图表内容
     * @param outputPath 输出路径
     * @throws ConversionException 转换失败
     */
    private void convertMermaidDiagram(String content, Path outputPath) throws ConversionException {
        if (!toolChecker.isToolAvailable("mmdc")) {
            throw new ConversionException("Mermaid CLI不可用，无法转换Mermaid图表");
        }
        
        try {
            // 创建临时输入文件
            Path tempInputPath = Files.createTempFile("mermaid_", ".mmd");
            
            try {
                // 写入Mermaid内容
                Files.writeString(tempInputPath, content, StandardCharsets.UTF_8);
                
                // 执行mmdc命令
                List<String> mmdcArgs = Arrays.asList(
                        "-i", tempInputPath.toString(),
                        "-o", outputPath.toString(),
                        "-f", "png",
                        "-b", "white",
                        "-s", "2" // 缩放因子
                );
                
                ProcessExecutor.ExecutionResult result = processExecutor.execute(
                        Arrays.asList("mmdc"), null, null);
                
                if (!result.isSuccess()) {
                    throw new ConversionException("Mermaid转换失败: " + result.getStderr());
                }
                
            } finally {
                // 清理临时文件
                try {
                    Files.deleteIfExists(tempInputPath);
                } catch (IOException e) {
                    logger.warn("清理临时Mermaid文件失败: {}", tempInputPath, e);
                }
            }
            
        } catch (IOException e) {
            throw new ConversionException("Mermaid转换失败", e);
        }
    }
    
    /**
     * 使用Batik将SVG转换为PNG
     * 
     * @param svgPath SVG文件路径
     * @param pngPath PNG输出路径
     * @throws ConversionException 转换失败
     */
    private void convertSvgToPngWithBatik(Path svgPath, Path pngPath) throws ConversionException {
        try {
            PNGTranscoder transcoder = new PNGTranscoder();
            
            // 设置转换参数
            float width = config.getConverterConfig("png_width", 800.0f);
            float height = config.getConverterConfig("png_height", 600.0f);
            
            transcoder.addTranscodingHint(PNGTranscoder.KEY_WIDTH, width);
            transcoder.addTranscodingHint(PNGTranscoder.KEY_HEIGHT, height);
            transcoder.addTranscodingHint(PNGTranscoder.KEY_BACKGROUND_COLOR, java.awt.Color.WHITE);
            
            // 执行转换
            try (InputStream svgInputStream = Files.newInputStream(svgPath);
                 OutputStream pngOutputStream = Files.newOutputStream(pngPath)) {
                
                TranscoderInput input = new TranscoderInput(svgInputStream);
                TranscoderOutput output = new TranscoderOutput(pngOutputStream);
                
                transcoder.transcode(input, output);
            }
            
        } catch (Exception e) {
            throw new ConversionException("Batik SVG转PNG失败", e);
        }
    }
    
    /**
     * 图表类型枚举
     */
    private enum DiagramType {
        PLANTUML,
        MERMAID,
        UNKNOWN
    }
}