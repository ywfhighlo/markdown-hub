package com.markdownhub.config;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import ch.qos.logback.classic.LoggerContext;
import ch.qos.logback.classic.encoder.PatternLayoutEncoder;
import ch.qos.logback.classic.spi.ILoggingEvent;
import ch.qos.logback.core.ConsoleAppender;
import ch.qos.logback.core.FileAppender;
import org.slf4j.LoggerFactory;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * 日志配置管理类
 * 负责动态配置应用程序的日志系统
 */
public class LoggingConfig {
    
    private static final String DEFAULT_PATTERN = "%d{yyyy-MM-dd HH:mm:ss} [%thread] %-5level %logger{36} - %msg%n";
    private static final String CONSOLE_PATTERN = "%d{HH:mm:ss} %-5level %logger{20} - %msg%n";
    
    /**
     * 配置日志系统
     * 
     * @param verbose 是否启用详细日志
     * @param logFile 日志文件路径（可选）
     */
    public static void configure(boolean verbose, String logFile) {
        LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();
        context.reset();
        
        // 配置控制台输出
        configureConsoleAppender(context, verbose);
        
        // 配置文件输出（如果指定）
        if (logFile != null && !logFile.trim().isEmpty()) {
            configureFileAppender(context, logFile);
        }
        
        // 配置根日志级别
        Logger rootLogger = context.getLogger(Logger.ROOT_LOGGER_NAME);
        rootLogger.setLevel(verbose ? Level.DEBUG : Level.INFO);
        
        // 配置第三方库日志级别
        configureThirdPartyLoggers(context, verbose);
    }
    
    /**
     * 配置控制台日志输出
     * 
     * @param context 日志上下文
     * @param verbose 是否详细模式
     */
    private static void configureConsoleAppender(LoggerContext context, boolean verbose) {
        ConsoleAppender<ILoggingEvent> consoleAppender = new ConsoleAppender<>();
        consoleAppender.setContext(context);
        consoleAppender.setName("CONSOLE");
        
        PatternLayoutEncoder encoder = new PatternLayoutEncoder();
        encoder.setContext(context);
        encoder.setPattern(verbose ? DEFAULT_PATTERN : CONSOLE_PATTERN);
        encoder.start();
        
        consoleAppender.setEncoder(encoder);
        consoleAppender.start();
        
        Logger rootLogger = context.getLogger(Logger.ROOT_LOGGER_NAME);
        rootLogger.addAppender(consoleAppender);
    }
    
    /**
     * 配置文件日志输出
     * 
     * @param context 日志上下文
     * @param logFile 日志文件路径
     */
    private static void configureFileAppender(LoggerContext context, String logFile) {
        try {
            Path logPath = Paths.get(logFile);
            Path parentDir = logPath.getParent();
            if (parentDir != null && !Files.exists(parentDir)) {
                Files.createDirectories(parentDir);
            }
            
            FileAppender<ILoggingEvent> fileAppender = new FileAppender<>();
            fileAppender.setContext(context);
            fileAppender.setName("FILE");
            fileAppender.setFile(logFile);
            fileAppender.setAppend(true);
            
            PatternLayoutEncoder encoder = new PatternLayoutEncoder();
            encoder.setContext(context);
            encoder.setPattern(DEFAULT_PATTERN);
            encoder.start();
            
            fileAppender.setEncoder(encoder);
            fileAppender.start();
            
            Logger rootLogger = context.getLogger(Logger.ROOT_LOGGER_NAME);
            rootLogger.addAppender(fileAppender);
            
        } catch (Exception e) {
            System.err.println("无法创建日志文件: " + logFile + ", 错误: " + e.getMessage());
        }
    }
    
    /**
     * 配置第三方库的日志级别
     * 
     * @param context 日志上下文
     * @param verbose 是否详细模式
     */
    private static void configureThirdPartyLoggers(LoggerContext context, boolean verbose) {
        // Apache POI 日志级别
        context.getLogger("org.apache.poi").setLevel(verbose ? Level.DEBUG : Level.WARN);
        
        // Jackson 日志级别
        context.getLogger("com.fasterxml.jackson").setLevel(verbose ? Level.DEBUG : Level.WARN);
        
        // Flexmark 日志级别
        context.getLogger("com.vladsch.flexmark").setLevel(verbose ? Level.DEBUG : Level.INFO);
        
        // Batik 日志级别
        context.getLogger("org.apache.batik").setLevel(verbose ? Level.DEBUG : Level.WARN);
        
        // PicoCLI 日志级别
        context.getLogger("picocli").setLevel(verbose ? Level.DEBUG : Level.INFO);
    }
    
    /**
     * 设置特定包的日志级别
     * 
     * @param packageName 包名
     * @param level 日志级别
     */
    public static void setLogLevel(String packageName, String level) {
        LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();
        Logger logger = context.getLogger(packageName);
        
        try {
            Level logLevel = Level.valueOf(level.toUpperCase());
            logger.setLevel(logLevel);
        } catch (IllegalArgumentException e) {
            System.err.println("无效的日志级别: " + level + ", 支持的级别: TRACE, DEBUG, INFO, WARN, ERROR");
        }
    }
    
    /**
     * 获取当前日志级别
     * 
     * @param packageName 包名
     * @return 日志级别字符串
     */
    public static String getLogLevel(String packageName) {
        LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();
        Logger logger = context.getLogger(packageName);
        Level level = logger.getLevel();
        return level != null ? level.toString() : "INHERITED";
    }
    
    /**
     * 启用性能日志
     * 用于记录转换操作的性能指标
     */
    public static void enablePerformanceLogging() {
        LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();
        context.getLogger("com.markdownhub.performance").setLevel(Level.DEBUG);
    }
    
    /**
     * 禁用性能日志
     */
    public static void disablePerformanceLogging() {
        LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();
        context.getLogger("com.markdownhub.performance").setLevel(Level.INFO);
    }
    
    /**
     * 创建性能日志记录器
     * 
     * @param clazz 调用类
     * @return 性能日志记录器
     */
    public static org.slf4j.Logger getPerformanceLogger(Class<?> clazz) {
        return LoggerFactory.getLogger("com.markdownhub.performance." + clazz.getSimpleName());
    }
    
    /**
     * 记录操作耗时
     * 
     * @param logger 日志记录器
     * @param operation 操作名称
     * @param startTime 开始时间（毫秒）
     */
    public static void logDuration(org.slf4j.Logger logger, String operation, long startTime) {
        long duration = System.currentTimeMillis() - startTime;
        if (logger.isDebugEnabled()) {
            logger.debug("{} 耗时: {}ms", operation, duration);
        }
    }
    
    /**
     * 记录操作耗时（纳秒精度）
     * 
     * @param logger 日志记录器
     * @param operation 操作名称
     * @param startNanos 开始时间（纳秒）
     */
    public static void logDurationNanos(org.slf4j.Logger logger, String operation, long startNanos) {
        long durationNanos = System.nanoTime() - startNanos;
        double durationMs = durationNanos / 1_000_000.0;
        if (logger.isDebugEnabled()) {
            logger.debug("{} 耗时: {:.2f}ms", operation, durationMs);
        }
    }
}