package com.markdownhub.external;

import com.markdownhub.config.Configuration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * 进程执行器
 * 负责执行外部命令和工具
 */
public class ProcessExecutor {
    
    private static final Logger logger = LoggerFactory.getLogger(ProcessExecutor.class);
    private static final Logger perfLogger = LoggerFactory.getLogger("com.markdownhub.performance.ProcessExecutor");
    
    private final Configuration config;
    
    public ProcessExecutor(Configuration config) {
        this.config = config;
    }
    
    /**
     * 执行命令结果
     */
    public static class ExecutionResult {
        private final int exitCode;
        private final String stdout;
        private final String stderr;
        private final long executionTimeMs;
        
        public ExecutionResult(int exitCode, String stdout, String stderr, long executionTimeMs) {
            this.exitCode = exitCode;
            this.stdout = stdout;
            this.stderr = stderr;
            this.executionTimeMs = executionTimeMs;
        }
        
        public int getExitCode() { return exitCode; }
        public String getStdout() { return stdout; }
        public String getStderr() { return stderr; }
        public long getExecutionTimeMs() { return executionTimeMs; }
        public boolean isSuccess() { return exitCode == 0; }
        
        @Override
        public String toString() {
            return String.format("ExecutionResult{exitCode=%d, executionTime=%dms, stdout='%s', stderr='%s'}",
                    exitCode, executionTimeMs, 
                    stdout.length() > 100 ? stdout.substring(0, 100) + "..." : stdout,
                    stderr.length() > 100 ? stderr.substring(0, 100) + "..." : stderr);
        }
    }
    
    /**
     * 执行命令
     * 
     * @param command 命令列表
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult execute(List<String> command) throws IOException {
        return execute(command, null, null);
    }
    
    /**
     * 执行命令
     * 
     * @param command 命令列表
     * @param workingDirectory 工作目录
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult execute(List<String> command, Path workingDirectory) throws IOException {
        return execute(command, workingDirectory, null);
    }
    
    /**
     * 执行命令
     * 
     * @param command 命令列表
     * @param workingDirectory 工作目录
     * @param input 标准输入内容
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult execute(List<String> command, Path workingDirectory, String input) throws IOException {
        long startTime = System.currentTimeMillis();
        
        if (command == null || command.isEmpty()) {
            throw new IllegalArgumentException("命令不能为空");
        }
        
        logger.debug("执行命令: {} (工作目录: {})", String.join(" ", command), workingDirectory);
        
        ProcessBuilder processBuilder = new ProcessBuilder(command);
        
        // 设置工作目录
        if (workingDirectory != null) {
            processBuilder.directory(workingDirectory.toFile());
        }
        
        // 合并错误流到标准输出
        processBuilder.redirectErrorStream(false);
        
        Process process = null;
        try {
            process = processBuilder.start();
            
            // 处理标准输入
            if (input != null && !input.isEmpty()) {
                try (OutputStreamWriter writer = new OutputStreamWriter(process.getOutputStream(), StandardCharsets.UTF_8)) {
                    writer.write(input);
                    writer.flush();
                }
            }
            
            // 读取输出
            String stdout = readStream(process.getInputStream());
            String stderr = readStream(process.getErrorStream());
            
            // 等待进程完成
            boolean finished = process.waitFor(config.getTimeoutSeconds(), TimeUnit.SECONDS);
            
            int exitCode;
            if (finished) {
                exitCode = process.exitValue();
            } else {
                logger.warn("命令执行超时 ({}秒): {}", config.getTimeoutSeconds(), String.join(" ", command));
                process.destroyForcibly();
                exitCode = -1;
                stderr = "命令执行超时";
            }
            
            long executionTime = System.currentTimeMillis() - startTime;
            ExecutionResult result = new ExecutionResult(exitCode, stdout, stderr, executionTime);
            
            if (perfLogger.isDebugEnabled()) {
                perfLogger.debug("命令执行完成: {} (耗时: {}ms, 退出码: {})", 
                        String.join(" ", command), executionTime, exitCode);
            }
            
            if (exitCode != 0) {
                logger.warn("命令执行失败: {} (退出码: {}, stderr: {})", 
                        String.join(" ", command), exitCode, stderr);
            }
            
            return result;
            
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("命令执行被中断: " + String.join(" ", command), e);
        } finally {
            if (process != null && process.isAlive()) {
                process.destroyForcibly();
            }
        }
    }
    
    /**
     * 读取输入流内容
     * 
     * @param inputStream 输入流
     * @return 流内容
     * @throws IOException 读取失败
     */
    private String readStream(InputStream inputStream) throws IOException {
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append(System.lineSeparator());
            }
        }
        return output.toString();
    }
    
    /**
     * 执行Java命令
     * 
     * @param jarPath JAR文件路径
     * @param args 命令参数
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult executeJava(String jarPath, List<String> args) throws IOException {
        List<String> command = new ArrayList<>();
        command.add(config.getJavaBinary());
        command.add("-jar");
        command.add(jarPath);
        command.addAll(args);
        
        return execute(command);
    }
    
    /**
     * 执行Java命令（带类路径）
     * 
     * @param classPath 类路径
     * @param mainClass 主类
     * @param args 命令参数
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult executeJavaWithClasspath(String classPath, String mainClass, List<String> args) throws IOException {
        List<String> command = new ArrayList<>();
        command.add(config.getJavaBinary());
        command.add("-cp");
        command.add(classPath);
        command.add(mainClass);
        command.addAll(args);
        
        return execute(command);
    }
    
    /**
     * 执行Pandoc命令
     * 
     * @param args Pandoc参数
     * @return 执行结果
     * @throws IOException 执行失败
     */
    public ExecutionResult executePandoc(List<String> args) throws IOException {
        List<String> command = new ArrayList<>();
        command.add(config.getPandocPath());
        command.addAll(args);
        
        return execute(command);
    }
    
    /**
     * 检查命令是否可用
     * 
     * @param command 命令名称
     * @return 是否可用
     */
    public boolean isCommandAvailable(String command) {
        try {
            List<String> testCommand = new ArrayList<>();
            
            // 根据操作系统选择测试命令
            String os = System.getProperty("os.name").toLowerCase();
            if (os.contains("win")) {
                testCommand.add("where");
            } else {
                testCommand.add("which");
            }
            testCommand.add(command);
            
            ExecutionResult result = execute(testCommand);
            return result.isSuccess();
            
        } catch (IOException e) {
            logger.debug("检查命令可用性失败: {}", command, e);
            return false;
        }
    }
    
    /**
     * 检查文件是否存在且可执行
     * 
     * @param filePath 文件路径
     * @return 是否存在且可执行
     */
    public boolean isFileExecutable(String filePath) {
        try {
            Path path = Paths.get(filePath);
            return Files.exists(path) && Files.isExecutable(path);
        } catch (Exception e) {
            logger.debug("检查文件可执行性失败: {}", filePath, e);
            return false;
        }
    }
    
    /**
     * 检查JAR文件是否存在
     * 
     * @param jarPath JAR文件路径
     * @return 是否存在
     */
    public boolean isJarFileAvailable(String jarPath) {
        try {
            Path path = Paths.get(jarPath);
            return Files.exists(path) && Files.isRegularFile(path) && jarPath.toLowerCase().endsWith(".jar");
        } catch (Exception e) {
            logger.debug("检查JAR文件失败: {}", jarPath, e);
            return false;
        }
    }
    
    /**
     * 获取命令版本信息
     * 
     * @param command 命令名称
     * @param versionArg 版本参数（如 --version）
     * @return 版本信息，如果获取失败返回null
     */
    public String getCommandVersion(String command, String versionArg) {
        try {
            List<String> versionCommand = new ArrayList<>();
            versionCommand.add(command);
            versionCommand.add(versionArg);
            
            ExecutionResult result = execute(versionCommand);
            if (result.isSuccess()) {
                String output = result.getStdout().trim();
                if (output.isEmpty()) {
                    output = result.getStderr().trim();
                }
                return output.split("\n")[0]; // 返回第一行
            }
        } catch (IOException e) {
            logger.debug("获取命令版本失败: {}", command, e);
        }
        return null;
    }
}