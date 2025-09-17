package com.markdownhub.external;

import com.markdownhub.config.Configuration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

/**
 * 系统工具检查器
 * 负责检查外部依赖工具的可用性
 */
public class SystemToolChecker {
    
    private static final Logger logger = LoggerFactory.getLogger(SystemToolChecker.class);
    
    private final Configuration config;
    private final ProcessExecutor processExecutor;
    private final Map<String, ToolStatus> toolStatusCache;
    
    public SystemToolChecker(Configuration config) {
        this.config = config;
        this.processExecutor = new ProcessExecutor(config);
        this.toolStatusCache = new HashMap<>();
    }
    
    /**
     * 工具状态
     */
    public static class ToolStatus {
        private final String name;
        private final boolean available;
        private final String version;
        private final String path;
        private final String errorMessage;
        
        public ToolStatus(String name, boolean available, String version, String path, String errorMessage) {
            this.name = name;
            this.available = available;
            this.version = version;
            this.path = path;
            this.errorMessage = errorMessage;
        }
        
        public String getName() { return name; }
        public boolean isAvailable() { return available; }
        public String getVersion() { return version; }
        public String getPath() { return path; }
        public String getErrorMessage() { return errorMessage; }
        
        @Override
        public String toString() {
            if (available) {
                return String.format("%s: 可用 (版本: %s, 路径: %s)", name, version != null ? version : "未知", path != null ? path : "系统PATH");
            } else {
                return String.format("%s: 不可用 (%s)", name, errorMessage != null ? errorMessage : "未找到");
            }
        }
    }
    
    /**
     * 检查所有必需的工具
     * 
     * @return 工具状态映射
     */
    public Map<String, ToolStatus> checkAllTools() {
        logger.info("开始检查系统工具...");
        
        Map<String, ToolStatus> results = new LinkedHashMap<>();
        
        // 检查Java
        results.put("java", checkJava());
        
        // 检查PlantUML JAR
        results.put("plantuml", checkPlantUmlJar());
        
        // 检查Batik库
        results.put("batik", checkBatikLib());
        
        // 检查Pandoc
        results.put("pandoc", checkPandoc());
        
        // 检查mmdc (Mermaid CLI)
        results.put("mmdc", checkMermaidCli());
        
        // 缓存结果
        toolStatusCache.putAll(results);
        
        // 输出检查结果摘要
        logToolCheckSummary(results);
        
        return results;
    }
    
    /**
     * 检查Java环境
     * 
     * @return Java工具状态
     */
    public ToolStatus checkJava() {
        try {
            String javaBinary = config.getJavaBinary();
            
            if (!processExecutor.isCommandAvailable(javaBinary)) {
                return new ToolStatus("Java", false, null, javaBinary, "Java命令不可用");
            }
            
            String version = processExecutor.getCommandVersion(javaBinary, "-version");
            return new ToolStatus("Java", true, version, javaBinary, null);
            
        } catch (Exception e) {
            logger.debug("检查Java失败", e);
            return new ToolStatus("Java", false, null, config.getJavaBinary(), e.getMessage());
        }
    }
    
    /**
     * 检查PlantUML JAR文件
     * 
     * @return PlantUML工具状态
     */
    public ToolStatus checkPlantUmlJar() {
        try {
            String jarPath = config.getPlantUmlJarPath();
            
            if (!processExecutor.isJarFileAvailable(jarPath)) {
                return new ToolStatus("PlantUML", false, null, jarPath, "PlantUML JAR文件不存在");
            }
            
            // 尝试获取PlantUML版本
            try {
                ProcessExecutor.ExecutionResult result = processExecutor.executeJava(jarPath, Arrays.asList("-version"));
                String version = null;
                if (result.isSuccess()) {
                    String output = result.getStdout().trim();
                    if (output.isEmpty()) {
                        output = result.getStderr().trim();
                    }
                    if (!output.isEmpty()) {
                        version = output.split("\n")[0];
                    }
                }
                return new ToolStatus("PlantUML", true, version, jarPath, null);
            } catch (IOException e) {
                // JAR文件存在但无法执行
                return new ToolStatus("PlantUML", false, null, jarPath, "无法执行PlantUML JAR: " + e.getMessage());
            }
            
        } catch (Exception e) {
            logger.debug("检查PlantUML失败", e);
            return new ToolStatus("PlantUML", false, null, config.getPlantUmlJarPath(), e.getMessage());
        }
    }
    
    /**
     * 检查Batik库
     * 
     * @return Batik工具状态
     */
    public ToolStatus checkBatikLib() {
        try {
            String batikLibPath = config.getBatikLibPath();
            Path libPath = Paths.get(batikLibPath);
            
            if (!Files.exists(libPath)) {
                return new ToolStatus("Batik", false, null, batikLibPath, "Batik库目录不存在");
            }
            
            if (!Files.isDirectory(libPath)) {
                return new ToolStatus("Batik", false, null, batikLibPath, "Batik库路径不是目录");
            }
            
            // 检查关键的Batik JAR文件
            String[] requiredJars = {
                "batik-transcoder.jar",
                "batik-codec.jar",
                "batik-dom.jar",
                "batik-svg-dom.jar"
            };
            
            List<String> missingJars = new ArrayList<>();
            for (String jarName : requiredJars) {
                Path jarPath = libPath.resolve(jarName);
                if (!Files.exists(jarPath)) {
                    missingJars.add(jarName);
                }
            }
            
            if (!missingJars.isEmpty()) {
                return new ToolStatus("Batik", false, null, batikLibPath, 
                        "缺少必需的JAR文件: " + String.join(", ", missingJars));
            }
            
            return new ToolStatus("Batik", true, "检测到必需的JAR文件", batikLibPath, null);
            
        } catch (Exception e) {
            logger.debug("检查Batik失败", e);
            return new ToolStatus("Batik", false, null, config.getBatikLibPath(), e.getMessage());
        }
    }
    
    /**
     * 检查Pandoc
     * 
     * @return Pandoc工具状态
     */
    public ToolStatus checkPandoc() {
        try {
            String pandocPath = config.getPandocPath();
            
            if (!processExecutor.isCommandAvailable(pandocPath)) {
                return new ToolStatus("Pandoc", false, null, pandocPath, "Pandoc命令不可用");
            }
            
            String version = processExecutor.getCommandVersion(pandocPath, "--version");
            return new ToolStatus("Pandoc", true, version, pandocPath, null);
            
        } catch (Exception e) {
            logger.debug("检查Pandoc失败", e);
            return new ToolStatus("Pandoc", false, null, config.getPandocPath(), e.getMessage());
        }
    }
    
    /**
     * 检查Mermaid CLI (mmdc)
     * 
     * @return Mermaid CLI工具状态
     */
    public ToolStatus checkMermaidCli() {
        try {
            String mmdcCommand = "mmdc";
            
            if (!processExecutor.isCommandAvailable(mmdcCommand)) {
                return new ToolStatus("Mermaid CLI", false, null, mmdcCommand, "mmdc命令不可用");
            }
            
            String version = processExecutor.getCommandVersion(mmdcCommand, "--version");
            return new ToolStatus("Mermaid CLI", true, version, mmdcCommand, null);
            
        } catch (Exception e) {
            logger.debug("检查Mermaid CLI失败", e);
            return new ToolStatus("Mermaid CLI", false, null, "mmdc", e.getMessage());
        }
    }
    
    /**
     * 检查特定工具是否可用
     * 
     * @param toolName 工具名称
     * @return 是否可用
     */
    public boolean isToolAvailable(String toolName) {
        ToolStatus status = toolStatusCache.get(toolName);
        if (status == null) {
            // 如果缓存中没有，执行检查
            switch (toolName.toLowerCase()) {
                case "java":
                    status = checkJava();
                    break;
                case "plantuml":
                    status = checkPlantUmlJar();
                    break;
                case "batik":
                    status = checkBatikLib();
                    break;
                case "pandoc":
                    status = checkPandoc();
                    break;
                case "mmdc":
                case "mermaid":
                    status = checkMermaidCli();
                    break;
                default:
                    return false;
            }
            toolStatusCache.put(toolName, status);
        }
        return status.isAvailable();
    }
    
    /**
     * 获取工具状态
     * 
     * @param toolName 工具名称
     * @return 工具状态
     */
    public ToolStatus getToolStatus(String toolName) {
        return toolStatusCache.get(toolName);
    }
    
    /**
     * 清除工具状态缓存
     */
    public void clearCache() {
        toolStatusCache.clear();
    }
    
    /**
     * 获取缺失的必需工具列表
     * 
     * @return 缺失工具列表
     */
    public List<String> getMissingRequiredTools() {
        List<String> missing = new ArrayList<>();
        
        // Java是必需的
        if (!isToolAvailable("java")) {
            missing.add("Java");
        }
        
        return missing;
    }
    
    /**
     * 获取缺失的可选工具列表
     * 
     * @return 缺失工具列表
     */
    public List<String> getMissingOptionalTools() {
        List<String> missing = new ArrayList<>();
        
        if (!isToolAvailable("plantuml")) {
            missing.add("PlantUML");
        }
        
        if (!isToolAvailable("batik")) {
            missing.add("Batik");
        }
        
        if (!isToolAvailable("pandoc")) {
            missing.add("Pandoc");
        }
        
        if (!isToolAvailable("mmdc")) {
            missing.add("Mermaid CLI");
        }
        
        return missing;
    }
    
    /**
     * 检查是否满足最低要求
     * 
     * @return 是否满足最低要求
     */
    public boolean meetsMinimumRequirements() {
        return getMissingRequiredTools().isEmpty();
    }
    
    /**
     * 输出工具检查结果摘要
     * 
     * @param toolStatuses 工具状态映射
     */
    private void logToolCheckSummary(Map<String, ToolStatus> toolStatuses) {
        logger.info("=== 系统工具检查结果 ===");
        
        int availableCount = 0;
        int totalCount = toolStatuses.size();
        
        for (ToolStatus status : toolStatuses.values()) {
            if (status.isAvailable()) {
                logger.info("✓ {}", status);
                availableCount++;
            } else {
                logger.warn("✗ {}", status);
            }
        }
        
        logger.info("工具可用性: {}/{}", availableCount, totalCount);
        
        List<String> missingRequired = getMissingRequiredTools();
        List<String> missingOptional = getMissingOptionalTools();
        
        if (!missingRequired.isEmpty()) {
            logger.error("缺少必需工具: {}", String.join(", ", missingRequired));
        }
        
        if (!missingOptional.isEmpty()) {
            logger.warn("缺少可选工具: {} (某些功能可能不可用)", String.join(", ", missingOptional));
        }
        
        if (meetsMinimumRequirements()) {
            logger.info("✓ 满足最低运行要求");
        } else {
            logger.error("✗ 不满足最低运行要求，请安装缺少的必需工具");
        }
        
        logger.info("========================");
    }
}