package com.markdownhub.config;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.PropertyNamingStrategies;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * 配置管理类
 * 负责管理应用程序的各种配置参数
 */
public class Configuration {
    
    private static final Logger logger = LoggerFactory.getLogger(Configuration.class);
    
    // 基本配置
    private String inputPath;
    private String outputPath;
    private String converterType;
    private String outputFormat;
    private boolean batchMode;
    private boolean verbose;
    private String configFile;
    
    // 转换器特定配置
    private Map<String, Object> converterConfig;
    
    // 工具路径配置
    private String plantUmlJarPath;
    private String batikLibPath;
    private String pandocPath;
    private String javaBinary;
    
    // 性能配置
    private int timeoutSeconds;
    private int maxConcurrentJobs;
    private long maxFileSize;
    
    // 私有构造函数，使用Builder模式
    private Configuration() {
        this.converterConfig = new HashMap<>();
        this.timeoutSeconds = 60;
        this.maxConcurrentJobs = 4;
        this.maxFileSize = 100 * 1024 * 1024; // 100MB
        this.plantUmlJarPath = "tools/plantuml.jar";
        this.batikLibPath = "tools/batik-lib";
        this.pandocPath = "pandoc";
        this.javaBinary = "java";
    }
    
    /**
     * 创建配置构建器
     * 
     * @return 配置构建器
     */
    public static Builder builder() {
        return new Builder();
    }
    
    /**
     * 从JSON文件加载配置
     * 
     * @param configFilePath 配置文件路径
     * @return 配置对象
     * @throws IOException 读取文件失败
     */
    public static Configuration fromFile(String configFilePath) throws IOException {
        Path configPath = Paths.get(configFilePath);
        if (!Files.exists(configPath)) {
            throw new IOException("配置文件不存在: " + configFilePath);
        }
        
        ObjectMapper mapper = new ObjectMapper();
        mapper.setPropertyNamingStrategy(PropertyNamingStrategies.SNAKE_CASE);
        
        Configuration config = mapper.readValue(configPath.toFile(), Configuration.class);
        config.configFile = configFilePath;
        
        logger.info("从文件加载配置: {}", configFilePath);
        return config;
    }
    
    /**
     * 保存配置到JSON文件
     * 
     * @param configFilePath 配置文件路径
     * @throws IOException 写入文件失败
     */
    public void saveToFile(String configFilePath) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.setPropertyNamingStrategy(PropertyNamingStrategies.SNAKE_CASE);
        
        Path configPath = Paths.get(configFilePath);
        Path parent = configPath.getParent();
        if (parent != null && !Files.exists(parent)) {
            Files.createDirectories(parent);
        }
        
        mapper.writerWithDefaultPrettyPrinter().writeValue(configPath.toFile(), this);
        logger.info("配置已保存到文件: {}", configFilePath);
    }
    
    // Getter方法
    public String getInputPath() { return inputPath; }
    public String getOutputPath() { return outputPath; }
    public String getConverterType() { return converterType; }
    public String getOutputFormat() { return outputFormat; }
    public boolean isBatchMode() { return batchMode; }
    public boolean isVerbose() { return verbose; }
    public String getConfigFile() { return configFile; }
    public Map<String, Object> getConverterConfig() { return converterConfig; }
    public String getPlantUmlJarPath() { return plantUmlJarPath; }
    public String getBatikLibPath() { return batikLibPath; }
    public String getPandocPath() { return pandocPath; }
    public String getJavaBinary() { return javaBinary; }
    public int getTimeoutSeconds() { return timeoutSeconds; }
    public int getMaxConcurrentJobs() { return maxConcurrentJobs; }
    public long getMaxFileSize() { return maxFileSize; }
    
    /**
     * 获取转换器特定配置
     * 
     * @param key 配置键
     * @param defaultValue 默认值
     * @param <T> 值类型
     * @return 配置值
     */
    @SuppressWarnings("unchecked")
    public <T> T getConverterConfig(String key, T defaultValue) {
        Object value = converterConfig.get(key);
        if (value == null) {
            return defaultValue;
        }
        try {
            return (T) value;
        } catch (ClassCastException e) {
            logger.warn("配置值类型不匹配: {} = {}, 使用默认值: {}", key, value, defaultValue);
            return defaultValue;
        }
    }
    
    /**
     * 设置转换器特定配置
     * 
     * @param key 配置键
     * @param value 配置值
     */
    public void setConverterConfig(String key, Object value) {
        converterConfig.put(key, value);
    }
    
    /**
     * 配置构建器
     */
    public static class Builder {
        private final Configuration config;
        
        private Builder() {
            this.config = new Configuration();
        }
        
        public Builder inputPath(String inputPath) {
            config.inputPath = inputPath;
            return this;
        }
        
        public Builder outputPath(String outputPath) {
            config.outputPath = outputPath;
            return this;
        }
        
        public Builder converterType(String converterType) {
            config.converterType = converterType;
            return this;
        }
        
        public Builder outputFormat(String outputFormat) {
            config.outputFormat = outputFormat;
            return this;
        }
        
        public Builder batchMode(boolean batchMode) {
            config.batchMode = batchMode;
            return this;
        }
        
        public Builder verbose(boolean verbose) {
            config.verbose = verbose;
            return this;
        }
        
        public Builder configFile(String configFile) {
            config.configFile = configFile;
            return this;
        }
        
        public Builder plantUmlJarPath(String path) {
            config.plantUmlJarPath = path;
            return this;
        }
        
        public Builder batikLibPath(String path) {
            config.batikLibPath = path;
            return this;
        }
        
        public Builder pandocPath(String path) {
            config.pandocPath = path;
            return this;
        }
        
        public Builder javaBinary(String path) {
            config.javaBinary = path;
            return this;
        }
        
        public Builder timeoutSeconds(int timeout) {
            config.timeoutSeconds = timeout;
            return this;
        }
        
        public Builder maxConcurrentJobs(int maxJobs) {
            config.maxConcurrentJobs = maxJobs;
            return this;
        }
        
        public Builder maxFileSize(long maxSize) {
            config.maxFileSize = maxSize;
            return this;
        }
        
        public Builder converterConfig(String key, Object value) {
            config.converterConfig.put(key, value);
            return this;
        }
        
        public Builder converterConfig(Map<String, Object> converterConfig) {
            config.converterConfig.putAll(converterConfig);
            return this;
        }
        
        /**
         * 从配置文件合并配置
         * 
         * @param configFilePath 配置文件路径
         * @return 构建器
         */
        public Builder mergeFromFile(String configFilePath) {
            try {
                Configuration fileConfig = Configuration.fromFile(configFilePath);
                
                // 只合并非空值
                if (fileConfig.plantUmlJarPath != null) config.plantUmlJarPath = fileConfig.plantUmlJarPath;
                if (fileConfig.batikLibPath != null) config.batikLibPath = fileConfig.batikLibPath;
                if (fileConfig.pandocPath != null) config.pandocPath = fileConfig.pandocPath;
                if (fileConfig.javaBinary != null) config.javaBinary = fileConfig.javaBinary;
                if (fileConfig.timeoutSeconds > 0) config.timeoutSeconds = fileConfig.timeoutSeconds;
                if (fileConfig.maxConcurrentJobs > 0) config.maxConcurrentJobs = fileConfig.maxConcurrentJobs;
                if (fileConfig.maxFileSize > 0) config.maxFileSize = fileConfig.maxFileSize;
                
                config.converterConfig.putAll(fileConfig.converterConfig);
                
            } catch (IOException e) {
                logger.warn("无法加载配置文件: {}, 使用默认配置", configFilePath, e);
            }
            return this;
        }
        
        public Configuration build() {
            // 如果指定了配置文件，尝试合并配置
            if (config.configFile != null && !config.configFile.isEmpty()) {
                mergeFromFile(config.configFile);
            }
            
            return config;
        }
    }
    
    @Override
    public String toString() {
        return "Configuration{" +
                "inputPath='" + inputPath + '\'' +
                ", outputPath='" + outputPath + '\'' +
                ", converterType='" + converterType + '\'' +
                ", outputFormat='" + outputFormat + '\'' +
                ", batchMode=" + batchMode +
                ", verbose=" + verbose +
                ", timeoutSeconds=" + timeoutSeconds +
                ", maxConcurrentJobs=" + maxConcurrentJobs +
                '}';
    }
}