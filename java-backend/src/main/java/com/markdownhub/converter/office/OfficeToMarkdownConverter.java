package com.markdownhub.converter.office;

import com.markdownhub.config.Configuration;
import com.markdownhub.converter.base.BaseConverter;
import com.markdownhub.converter.base.ConversionException;
import com.markdownhub.converter.base.DependencyException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

/**
 * Office到Markdown转换器
 * 支持将Word (.docx) 和PowerPoint (.pptx) 文档转换为Markdown格式
 */
public class OfficeToMarkdownConverter extends BaseConverter {
    
    private static final Logger logger = LoggerFactory.getLogger(OfficeToMarkdownConverter.class);
    private static final Logger perfLogger = LoggerFactory.getLogger("com.markdownhub.performance.OfficeToMarkdownConverter");
    
    public OfficeToMarkdownConverter(Configuration config) {
        super(config);
    }
    
    @Override
    public String getConverterType() {
        return "office-to-md";
    }
    
    @Override
    public List<String> getSupportedInputFormats() {
        return Arrays.asList("docx", "pptx");
    }
    
    @Override
    public List<String> getSupportedOutputFormats() {
        return Arrays.asList("md", "markdown");
    }
    
    @Override
    protected void checkDependencies() throws DependencyException {
        // Apache POI库已通过Maven依赖包含，无需额外检查
        logger.debug("Office到Markdown转换器依赖检查完成");
    }
    
    @Override
    protected void performConversion(Path inputPath, Path outputPath) throws ConversionException {
        try {
            long startTime = System.nanoTime();
            
            // 根据输入格式选择转换方法
            String inputFormat = getInputFormat(inputPath);
            String markdownContent;
            
            switch (inputFormat.toLowerCase()) {
                case "docx":
                    markdownContent = convertWordToMarkdown(inputPath);
                    break;
                case "pptx":
                    markdownContent = convertPowerPointToMarkdown(inputPath);
                    break;
                default:
                    throw new ConversionException("不支持的输入格式: " + inputFormat);
            }
            
            // 写入Markdown文件
            Files.writeString(outputPath, markdownContent, StandardCharsets.UTF_8);
            
            if (perfLogger.isDebugEnabled()) {
                long duration = System.nanoTime() - startTime;
                perfLogger.debug("Office到Markdown转换完成: {} -> {} (耗时: {:.2f}ms)", 
                        inputPath.getFileName(), outputPath.getFileName(), duration / 1_000_000.0);
            }
            
        } catch (IOException e) {
            throw new ConversionException("Office文档转换失败: " + inputPath, e);
        }
    }
    
    /**
     * 转换Word文档为Markdown
     * 
     * @param inputPath Word文档路径
     * @return Markdown内容
     * @throws ConversionException 转换失败
     */
    private String convertWordToMarkdown(Path inputPath) throws ConversionException {
        try {
            FileInputStream fis = new FileInputStream(inputPath.toFile());
            XWPFDocument document = new XWPFDocument(fis);
            
            StringBuilder markdown = new StringBuilder();
            WordToMarkdownConverter converter = new WordToMarkdownConverter(config);
            
            // 处理文档元素
            for (IBodyElement element : document.getBodyElements()) {
                converter.processElement(element, markdown);
            }
            
            document.close();
            fis.close();
            
            return markdown.toString();
            
        } catch (IOException e) {
            throw new ConversionException("读取Word文档失败: " + inputPath, e);
        }
    }
    
    /**
     * 转换PowerPoint文档为Markdown
     * 
     * @param inputPath PowerPoint文档路径
     * @return Markdown内容
     * @throws ConversionException 转换失败
     */
    private String convertPowerPointToMarkdown(Path inputPath) throws ConversionException {
        try (FileInputStream fis = new FileInputStream(inputPath.toFile());
             XMLSlideShow presentation = new XMLSlideShow(fis)) {
            
            StringBuilder markdown = new StringBuilder();
            PowerPointToMarkdownConverter converter = new PowerPointToMarkdownConverter(config);
            
            // 处理幻灯片
            List<XSLFSlide> slides = presentation.getSlides();
            for (int i = 0; i < slides.size(); i++) {
                converter.processSlide(slides.get(i), i + 1, markdown);
            }
            
            return markdown.toString();
            
        } catch (IOException e) {
            throw new ConversionException("读取PowerPoint文档失败: " + inputPath, e);
        }
    }
    
    /**
     * Word到Markdown转换器
     */
    private static class WordToMarkdownConverter {
        private final Configuration config;
        
        public WordToMarkdownConverter(Configuration config) {
            this.config = config;
        }
        
        /**
         * 处理文档元素
         * 
         * @param element 文档元素
         * @param markdown Markdown构建器
         */
        public void processElement(IBodyElement element, StringBuilder markdown) {
            if (element instanceof XWPFParagraph) {
                processParagraph((XWPFParagraph) element, markdown);
            } else if (element instanceof XWPFTable) {
                processTable((XWPFTable) element, markdown);
            }
        }
        
        /**
         * 处理段落
         * 
         * @param paragraph 段落
         * @param markdown Markdown构建器
         */
        private void processParagraph(XWPFParagraph paragraph, StringBuilder markdown) {
            String text = paragraph.getText().trim();
            if (text.isEmpty()) {
                markdown.append("\n");
                return;
            }
            
            // 检查是否为标题
            String style = paragraph.getStyle();
            if (style != null && style.startsWith("Heading")) {
                int level = extractHeadingLevel(style);
                markdown.append("#".repeat(Math.max(1, level))).append(" ").append(text).append("\n\n");
                return;
            }
            
            // 检查格式
            boolean isBold = false;
            boolean isItalic = false;
            boolean isCode = false;
            
            List<XWPFRun> runs = paragraph.getRuns();
            if (!runs.isEmpty()) {
                XWPFRun firstRun = runs.get(0);
                isBold = firstRun.isBold();
                isItalic = firstRun.isItalic();
                
                String fontFamily = firstRun.getFontFamily();
                isCode = fontFamily != null && (fontFamily.equals("Consolas") || fontFamily.equals("Courier New"));
            }
            
            // 检查是否为列表项
            if (paragraph.getNumID() != null) {
                markdown.append("- ").append(formatText(text, isBold, isItalic, isCode)).append("\n");
            } else if (isCode && text.contains("\n")) {
                // 代码块
                markdown.append("```\n").append(text).append("\n```\n\n");
            } else {
                // 普通段落
                markdown.append(formatText(text, isBold, isItalic, isCode)).append("\n\n");
            }
        }
        
        /**
         * 处理表格
         * 
         * @param table 表格
         * @param markdown Markdown构建器
         */
        private void processTable(XWPFTable table, StringBuilder markdown) {
            List<XWPFTableRow> rows = table.getRows();
            if (rows.isEmpty()) {
                return;
            }
            
            // 处理表头
            XWPFTableRow headerRow = rows.get(0);
            markdown.append("|");
            for (XWPFTableCell cell : headerRow.getTableCells()) {
                markdown.append(" ").append(cell.getText().trim()).append(" |");
            }
            markdown.append("\n");
            
            // 添加分隔行
            markdown.append("|");
            for (int i = 0; i < headerRow.getTableCells().size(); i++) {
                markdown.append(" --- |");
            }
            markdown.append("\n");
            
            // 处理数据行
            for (int i = 1; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);
                markdown.append("|");
                for (XWPFTableCell cell : row.getTableCells()) {
                    markdown.append(" ").append(cell.getText().trim()).append(" |");
                }
                markdown.append("\n");
            }
            
            markdown.append("\n");
        }
        
        /**
         * 提取标题级别
         * 
         * @param style 样式名称
         * @return 标题级别
         */
        private int extractHeadingLevel(String style) {
            try {
                if (style.equals("Heading1")) return 1;
                if (style.equals("Heading2")) return 2;
                if (style.equals("Heading3")) return 3;
                if (style.equals("Heading4")) return 4;
                if (style.equals("Heading5")) return 5;
                if (style.equals("Heading6")) return 6;
                
                // 尝试从样式名称中提取数字
                String number = style.replaceAll("[^0-9]", "");
                if (!number.isEmpty()) {
                    return Math.min(6, Integer.parseInt(number));
                }
            } catch (NumberFormatException e) {
                // 忽略解析错误
            }
            return 1;
        }
        
        /**
         * 格式化文本
         * 
         * @param text 原始文本
         * @param bold 是否粗体
         * @param italic 是否斜体
         * @param code 是否代码
         * @return 格式化后的文本
         */
        private String formatText(String text, boolean bold, boolean italic, boolean code) {
            if (code) {
                return "`" + text + "`";
            }
            
            String result = text;
            if (bold) {
                result = "**" + result + "**";
            }
            if (italic) {
                result = "*" + result + "*";
            }
            
            return result;
        }
    }
    
    /**
     * PowerPoint到Markdown转换器
     */
    private static class PowerPointToMarkdownConverter {
        private final Configuration config;
        
        public PowerPointToMarkdownConverter(Configuration config) {
            this.config = config;
        }
        
        /**
         * 处理幻灯片
         * 
         * @param slide 幻灯片
         * @param slideNumber 幻灯片编号
         * @param markdown Markdown构建器
         */
        public void processSlide(XSLFSlide slide, int slideNumber, StringBuilder markdown) {
            // 添加幻灯片标题
            markdown.append("# 幻灯片 ").append(slideNumber).append("\n\n");
            
            // 处理幻灯片中的形状
            for (XSLFShape shape : slide.getShapes()) {
                processShape(shape, markdown);
            }
            
            markdown.append("---\n\n"); // 幻灯片分隔符
        }
        
        /**
         * 处理形状
         * 
         * @param shape 形状
         * @param markdown Markdown构建器
         */
        private void processShape(XSLFShape shape, StringBuilder markdown) {
            if (shape instanceof XSLFTextShape) {
                processTextShape((XSLFTextShape) shape, markdown);
            } else if (shape instanceof XSLFGroupShape) {
                // 处理组合形状
                XSLFGroupShape group = (XSLFGroupShape) shape;
                for (XSLFShape childShape : group.getShapes()) {
                    processShape(childShape, markdown);
                }
            }
            // 其他类型的形状（如图片、图表）暂时跳过
        }
        
        /**
         * 处理文本形状
         * 
         * @param textShape 文本形状
         * @param markdown Markdown构建器
         */
        private void processTextShape(XSLFTextShape textShape, StringBuilder markdown) {
            List<XSLFTextParagraph> paragraphs = textShape.getTextParagraphs();
            
            for (XSLFTextParagraph paragraph : paragraphs) {
                String text = paragraph.getText().trim();
                if (text.isEmpty()) {
                    continue;
                }
                
                // 检查是否为标题（通常字体较大或加粗）
                boolean isTitle = false;
                boolean isBold = false;
                boolean isItalic = false;
                
                List<XSLFTextRun> runs = paragraph.getTextRuns();
                if (!runs.isEmpty()) {
                    XSLFTextRun firstRun = runs.get(0);
                    Double fontSize = firstRun.getFontSize();
                    isBold = firstRun.isBold();
                    isItalic = firstRun.isItalic();
                    
                    // 如果字体大小大于18或者加粗，认为是标题
                    isTitle = (fontSize != null && fontSize > 18) || isBold;
                }
                
                // 检查缩进级别
                Double indent = paragraph.getIndent();
                boolean isListItem = indent != null && indent > 0;
                
                if (isTitle) {
                    markdown.append("## ").append(text).append("\n\n");
                } else if (isListItem) {
                    markdown.append("- ").append(formatPowerPointText(text, isBold, isItalic)).append("\n");
                } else {
                    markdown.append(formatPowerPointText(text, isBold, isItalic)).append("\n\n");
                }
            }
        }
        
        /**
         * 格式化PowerPoint文本
         * 
         * @param text 原始文本
         * @param bold 是否粗体
         * @param italic 是否斜体
         * @return 格式化后的文本
         */
        private String formatPowerPointText(String text, boolean bold, boolean italic) {
            String result = text;
            
            if (bold) {
                result = "**" + result + "**";
            }
            if (italic) {
                result = "*" + result + "*";
            }
            
            return result;
        }
    }
}