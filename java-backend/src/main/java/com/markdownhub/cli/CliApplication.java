package com.markdownhub.cli;

import com.markdownhub.config.Configuration;
import com.markdownhub.converter.base.BaseConverter;
import com.markdownhub.converter.base.ConverterFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import picocli.CommandLine.Parameters;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.concurrent.Callable;

/**
 * Markdown Hub CLI应用程序主入口
 * 使用PicoCLI框架处理命令行参数解析和路由
 */
@Command(
    name = "markdown-hub",
    mixinStandardHelpOptions = true,
    version = "Markdown Hub 1.0.0",
    description = "Document conversion tool for Markdown, Office, and diagram formats"
)
public class CliApplication implements Callable<Integer> {
    
    private static final Logger logger = LoggerFactory.getLogger(CliApplication.class);
    
    @Parameters(
        index = "0",
        description = "转换类型: diagram-to-png, md-to-office, office-to-md, plantuml-to-png"
    )
    private String converterType;
    
    @Parameters(
        index = "1",
        description = "输入文件或目录路径"
    )
    private String inputPath;
    
    @Parameters(
        index = "2",
        description = "输出文件或目录路径"
    )
    private String outputPath;
    
    @Option(
        names = {"--format", "-f"},
        description = "输出格式 (docx, pptx, pdf, html, png等)"
    )
    private String format;
    
    @Option(
        names = {"--config", "-c"},
        description = "配置文件路径"
    )
    private String configPath;
    
    @Option(
        names = {"--verbose", "-v"},
        description = "启用详细日志输出"
    )
    private boolean verbose;
    
    @Option(
        names = {"--dry-run"},
        description = "仅显示将要执行的操作，不实际执行"
    )
    private boolean dryRun;
    
    @Option(
        names = {"--batch", "-b"},
        description = "批量处理模式"
    )
    private boolean batchMode;
    
    public static void main(String[] args) {
        int exitCode = new CommandLine(new CliApplication()).execute(args);
        System.exit(exitCode);
    }
    
    @Override
    public Integer call() throws Exception {
        try {
            // 配置日志级别
            configureLogging();
            
            // 验证输入参数
            if (!validateArguments()) {
                return 1;
            }
            
            // 加载配置
            Configuration config = loadConfiguration();
            
            // 创建转换器
            BaseConverter converter = ConverterFactory.createConverter(converterType, config);
            if (converter == null) {
                logger.error("不支持的转换类型: {}", converterType);
                return 1;
            }
            
            // 执行转换
            if (dryRun) {
                logger.info("[DRY RUN] 将执行转换: {} -> {}", inputPath, outputPath);
                logger.info("[DRY RUN] 转换类型: {}", converterType);
                if (format != null) {
                    logger.info("[DRY RUN] 输出格式: {}", format);
                }
                return 0;
            }
            
            logger.info("开始转换: {} -> {}", inputPath, outputPath);
            Path inputPathObj = Paths.get(inputPath);
            Path outputPathObj = Paths.get(outputPath);
            converter.convert(inputPathObj, outputPathObj);
            
            logger.info("转换完成: {} -> {}", inputPath, outputPath);
            
            return 0;
            
        } catch (Exception e) {
            logger.error("转换过程中发生错误: {}", e.getMessage(), e);
            return 1;
        }
    }
    
    private void configureLogging() {
        // 根据verbose参数配置日志级别
        if (verbose) {
            System.setProperty("logging.level.com.markdownhub", "DEBUG");
            logger.debug("启用详细日志输出");
        }
    }
    
    private boolean validateArguments() {
        // 验证输入路径
        Path input = Paths.get(inputPath);
        if (!Files.exists(input)) {
            logger.error("输入路径不存在: {}", inputPath);
            return false;
        }
        
        // 验证输出路径的父目录
        Path output = Paths.get(outputPath);
        Path outputParent = output.getParent();
        if (outputParent != null && !Files.exists(outputParent)) {
            try {
                Files.createDirectories(outputParent);
                logger.debug("创建输出目录: {}", outputParent);
            } catch (Exception e) {
                logger.error("无法创建输出目录: {}", outputParent, e);
                return false;
            }
        }
        
        // 验证转换器类型
        if (!isValidConverterType(converterType)) {
            logger.error("不支持的转换类型: {}. 支持的类型: diagram-to-png, md-to-office, office-to-md, plantuml-to-png", 
                        converterType);
            return false;
        }
        
        return true;
    }
    
    private boolean isValidConverterType(String type) {
        return type != null && (
            "diagram-to-png".equals(type) ||
            "md-to-office".equals(type) ||
            "office-to-md".equals(type) ||
            "plantuml-to-png".equals(type)
        );
    }
    
    private Configuration loadConfiguration() {
        Configuration.Builder builder = Configuration.builder()
            .inputPath(inputPath)
            .outputPath(outputPath)
            .converterType(converterType)
            .batchMode(batchMode)
            .verbose(verbose);
        
        if (format != null) {
            builder.outputFormat(format);
        }
        
        if (configPath != null) {
            builder.configFile(configPath);
        }
        
        return builder.build();
    }
}