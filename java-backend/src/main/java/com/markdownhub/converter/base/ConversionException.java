package com.markdownhub.converter.base;

/**
 * 转换异常类
 * 用于表示转换过程中发生的错误
 */
public class ConversionException extends Exception {
    
    /**
     * 构造函数
     * 
     * @param message 异常消息
     */
    public ConversionException(String message) {
        super(message);
    }
    
    /**
     * 构造函数
     * 
     * @param message 异常消息
     * @param cause 原因异常
     */
    public ConversionException(String message, Throwable cause) {
        super(message, cause);
    }
    
    /**
     * 构造函数
     * 
     * @param cause 原因异常
     */
    public ConversionException(Throwable cause) {
        super(cause);
    }
}