package com.markdownhub.converter.base;

/**
 * 依赖异常类
 * 用于表示转换器依赖检查失败的错误
 */
public class DependencyException extends Exception {
    
    /**
     * 构造函数
     * 
     * @param message 异常消息
     */
    public DependencyException(String message) {
        super(message);
    }
    
    /**
     * 构造函数
     * 
     * @param message 异常消息
     * @param cause 原因异常
     */
    public DependencyException(String message, Throwable cause) {
        super(message, cause);
    }
    
    /**
     * 构造函数
     * 
     * @param cause 原因异常
     */
    public DependencyException(Throwable cause) {
        super(cause);
    }
}