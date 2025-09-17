package com.markdownhub.converter.base;

import com.markdownhub.config.Configuration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 转换器基类
 * 定义所有转换器的通用接口和行为
 */
public abstract class BaseConverter {
    
    protected final Configuration config;
    protected final Logger logger;
    
    /**
     * 构造函数
     * 
     * @param config 配置对象
     */
    public BaseConverter(Configuration config) {
        this.config = config;
        this.logger = LoggerFactory.getLogger(this.getClass());
    }
    
    /**
     * 执行转换操作
     * 
     * @param inputPath 输入文件路径
     * @param outputPath 输出文件路径
     * @throws ConversionException 转换过程中发生错误
     */
    public final void convert(Path inputPath, Path outputPath) throws ConversionException {
        logger.info("开始转换: {} -> {}", inputPath, outputPath);
        
        try {
            // 检查依赖
            checkDependencies();
            
            // 验证输入
            if (!validateInput(inputPath)) {
                throw new ConversionException("输入验证失败: " + inputPath);
            }
            
            // 验证输出目录
            ensureOutputDirectory(outputPath.getParent());
            
            // 执行转换
            performConversion(inputPath, outputPath);
            
            // 验证输出
            if (!Files.exists(outputPath)) {
                throw new ConversionException("转换失败，输出文件未生成: " + outputPath);
            }
            
            logger.info("转换完成: {} -> {}", inputPath, outputPath);
            
        } catch (DependencyException e) {
            logger.error("依赖检查失败: {}", e.getMessage());
            throw new ConversionException("转换失败，依赖检查未通过", e);
        } catch (IOException e) {
            logger.error("IO操作失败: {}", e.getMessage(), e);
            throw new ConversionException("转换失败，IO错误: " + inputPath, e);
        } catch (Exception e) {
            logger.error("转换过程中发生错误: {}", e.getMessage(), e);
            if (e instanceof ConversionException) {
                throw e;
            } else {
                throw new ConversionException("转换失败: " + inputPath, e);
            }
        }
    }
    
    /**
     * 获取转换器类型标识
     * 
     * @return 转换器类型
     */
    public abstract String getConverterType();
    
    /**
     * 获取支持的输入格式
     * 
     * @return 支持的输入格式列表
     */
    public abstract List<String> getSupportedInputFormats();
    
    /**
     * 获取支持的输出格式
     * 
     * @return 支持的输出格式列表
     */
    public abstract List<String> getSupportedOutputFormats();
    
    /**
     * 执行具体的转换逻辑（由子类实现）
     * 
     * @param inputPath 输入文件路径
     * @param outputPath 输出文件路径
     * @throws ConversionException 转换异常
     */
    protected abstract void performConversion(Path inputPath, Path outputPath) throws ConversionException;
    
    /**
     * 检查转换器依赖（由子类实现）
     * 
     * @throws DependencyException 依赖检查失败
     */
    protected abstract void checkDependencies() throws DependencyException;
    
    /**
     * 验证输入文件
     * 
     * @param inputPath 输入路径
     * @return 验证结果
     */
    protected boolean validateInput(Path inputPath) {
        if (!Files.exists(inputPath)) {
            logger.error("输入文件不存在: {}", inputPath);
            return false;
        }
        
        if (!Files.isRegularFile(inputPath)) {
            logger.error("输入路径不是文件: {}", inputPath);
            return false;
        }
        
        return validateFileInput(inputPath);
    }
    
    /**
     * 获取输入文件格式
     * 
     * @param inputPath 输入文件路径
     * @return 文件格式（扩展名）
     */
    protected String getInputFormat(Path inputPath) {
        String fileName = inputPath.getFileName().toString();
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot > 0 && lastDot < fileName.length() - 1) {
            return fileName.substring(lastDot + 1).toLowerCase();
        }
        return "";
    }
    
    /**
     * 验证文件输入
     * 
     * @param file 文件路径
     * @return 验证结果
     */
    protected boolean validateFileInput(Path file) {
        if (!Files.isReadable(file)) {
            logger.error("文件不可读: {}", file);
            return false;
        }
        
        String format = getInputFormat(file);
        List<String> supportedFormats = getSupportedInputFormats();
        
        if (!supportedFormats.contains(format)) {
            logger.error("不支持的文件格式: {} (支持的格式: {})", format, supportedFormats);
            return false;
        }
        
        return true;
    }
    
    /**
     * 确保输出目录存在
     * 
     * @param outputDir 输出目录路径
     * @throws IOException 创建目录失败
     */
    protected void ensureOutputDirectory(Path outputDir) throws IOException {
        if (outputDir != null && !Files.exists(outputDir)) {
            Files.createDirectories(outputDir);
            logger.debug("创建输出目录: {}", outputDir);
        }
    }
    
    /**
     * 获取文件扩展名
     * 
     * @param fileName 文件名
     * @return 扩展名（不包含点号）
     */
    protected String getFileExtension(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot > 0 && lastDot < fileName.length() - 1) {
            return fileName.substring(lastDot + 1);
        }
        return "";
    }
    
    /**
     * 获取文件基本名称（不包含扩展名）
     * 
     * @param fileName 文件名
     * @return 基本名称
     */
    protected String getBaseName(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot > 0) {
            return fileName.substring(0, lastDot);
        }
        return fileName;
    }
    
    /**
     * 获取配置对象
     * 
     * @return 配置对象
     */
    public Configuration getConfig() {
        return config;
    }
}