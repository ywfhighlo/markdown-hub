package com.markdownhub.converter.base;

import com.markdownhub.config.Configuration;
import com.markdownhub.converter.diagram.DiagramToPngConverter;
import com.markdownhub.converter.office.MarkdownToOfficeConverter;
import com.markdownhub.converter.office.OfficeToMarkdownConverter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;

/**
 * 转换器工厂类
 * 负责创建和管理各种类型的转换器实例
 */
public class ConverterFactory {
    
    private static final Logger logger = LoggerFactory.getLogger(ConverterFactory.class);
    
    // 转换器类型到构造函数的映射
    private static final Map<String, Function<Configuration, BaseConverter>> CONVERTER_REGISTRY = new HashMap<>();
    
    static {
        // 注册所有支持的转换器
        registerConverter("diagram-to-png", DiagramToPngConverter::new);
        registerConverter("plantuml-to-png", DiagramToPngConverter::new);
        registerConverter("md-to-office", MarkdownToOfficeConverter::new);
        registerConverter("office-to-md", OfficeToMarkdownConverter::new);
    }
    
    /**
     * 注册转换器
     * 
     * @param type 转换器类型
     * @param constructor 转换器构造函数
     */
    public static void registerConverter(String type, Function<Configuration, BaseConverter> constructor) {
        CONVERTER_REGISTRY.put(type, constructor);
        logger.debug("注册转换器: {}", type);
    }
    
    /**
     * 创建转换器实例
     * 
     * @param type 转换器类型
     * @param config 配置对象
     * @return 转换器实例，如果类型不支持则返回null
     */
    public static BaseConverter createConverter(String type, Configuration config) {
        if (type == null || type.trim().isEmpty()) {
            logger.error("转换器类型不能为空");
            return null;
        }
        
        Function<Configuration, BaseConverter> constructor = CONVERTER_REGISTRY.get(type);
        if (constructor == null) {
            logger.error("不支持的转换器类型: {}. 支持的类型: {}", type, getSupportedTypes());
            return null;
        }
        
        try {
            BaseConverter converter = constructor.apply(config);
            logger.debug("创建转换器成功: {} -> {}", type, converter.getClass().getSimpleName());
            return converter;
        } catch (Exception e) {
            logger.error("创建转换器失败: {}", type, e);
            return null;
        }
    }
    
    /**
     * 获取所有支持的转换器类型
     * 
     * @return 支持的转换器类型列表
     */
    public static String[] getSupportedTypes() {
        return CONVERTER_REGISTRY.keySet().toArray(new String[0]);
    }
    
    /**
     * 检查是否支持指定的转换器类型
     * 
     * @param type 转换器类型
     * @return 如果支持返回true，否则返回false
     */
    public static boolean isSupported(String type) {
        return type != null && CONVERTER_REGISTRY.containsKey(type);
    }
    
    /**
     * 获取转换器的描述信息
     * 
     * @param type 转换器类型
     * @return 转换器描述，如果类型不支持则返回null
     */
    public static String getConverterDescription(String type) {
        switch (type) {
            case "diagram-to-png":
                return "将各种图表格式转换为PNG图片";
            case "plantuml-to-png":
                return "将PlantUML文件转换为PNG图片";
            case "md-to-office":
                return "将Markdown文件转换为Office文档";
            case "office-to-md":
                return "将Office文档转换为Markdown文件";
            default:
                return null;
        }
    }
    
    /**
     * 打印所有支持的转换器信息
     */
    public static void printSupportedConverters() {
        logger.info("支持的转换器类型:");
        for (String type : getSupportedTypes()) {
            String description = getConverterDescription(type);
            logger.info("  - {}: {}", type, description != null ? description : "无描述");
        }
    }
}