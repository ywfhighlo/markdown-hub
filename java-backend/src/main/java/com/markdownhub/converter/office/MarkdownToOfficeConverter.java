package com.markdownhub.converter.office;

import com.markdownhub.config.Configuration;
import com.markdownhub.converter.base.BaseConverter;
import com.markdownhub.converter.base.ConversionException;
import com.markdownhub.converter.base.DependencyException;
import com.vladsch.flexmark.ast.*;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.ext.gfm.strikethrough.StrikethroughExtension;
import com.vladsch.flexmark.ext.autolink.AutolinkExtension;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.util.Units;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

/**
 * Markdown到Office转换器
 * 支持将Markdown文档转换为Word (.docx) 和PowerPoint (.pptx) 格式
 */
public class MarkdownToOfficeConverter extends BaseConverter {
    
    private static final Logger logger = LoggerFactory.getLogger(MarkdownToOfficeConverter.class);
    private static final Logger perfLogger = LoggerFactory.getLogger("com.markdownhub.performance.MarkdownToOfficeConverter");
    
    private final Parser markdownParser;
    
    public MarkdownToOfficeConverter(Configuration config) {
        super(config);
        this.markdownParser = createMarkdownParser();
    }
    
    @Override
    public String getConverterType() {
        return "md-to-office";
    }
    
    @Override
    public List<String> getSupportedInputFormats() {
        return Arrays.asList("md", "markdown", "txt");
    }
    
    @Override
    public List<String> getSupportedOutputFormats() {
        return Arrays.asList("docx", "pptx");
    }
    
    @Override
    protected void checkDependencies() throws DependencyException {
        // Apache POI库已通过Maven依赖包含，无需额外检查
        logger.debug("Markdown到Office转换器依赖检查完成");
    }
    
    @Override
    protected void performConversion(Path inputPath, Path outputPath) throws ConversionException {
        try {
            long startTime = System.nanoTime();
            
            // 读取Markdown内容
            String markdownContent = Files.readString(inputPath, StandardCharsets.UTF_8);
            
            // 解析Markdown
            Node document = markdownParser.parse(markdownContent);
            
            // 根据输出格式选择转换方法
            String outputFormat = getInputFormat(outputPath);
            switch (outputFormat.toLowerCase()) {
                case "docx":
                    convertToWord(document, outputPath);
                    break;
                case "pptx":
                    convertToPowerPoint(document, outputPath);
                    break;
                default:
                    throw new ConversionException("不支持的输出格式: " + outputFormat);
            }
            
            if (perfLogger.isDebugEnabled()) {
                long duration = System.nanoTime() - startTime;
                perfLogger.debug("Markdown到Office转换完成: {} -> {} (耗时: {:.2f}ms)", 
                        inputPath.getFileName(), outputPath.getFileName(), duration / 1_000_000.0);
            }
            
        } catch (IOException e) {
            throw new ConversionException("读取Markdown文件失败: " + inputPath, e);
        }
    }
    
    /**
     * 创建Markdown解析器
     * 
     * @return Markdown解析器
     */
    private Parser createMarkdownParser() {
        MutableDataSet options = new MutableDataSet();
        
        // 启用扩展
        options.set(Parser.EXTENSIONS, Arrays.asList(
                TablesExtension.create(),
                StrikethroughExtension.create(),
                AutolinkExtension.create()
        ));
        
        return Parser.builder(options).build();
    }
    
    /**
     * 转换为Word文档
     * 
     * @param document Markdown AST文档
     * @param outputPath 输出路径
     * @throws ConversionException 转换失败
     */
    private void convertToWord(Node document, Path outputPath) throws ConversionException {
        try (XWPFDocument wordDoc = new XWPFDocument()) {
            
            WordDocumentBuilder builder = new WordDocumentBuilder(wordDoc, config);
            
            // 遍历Markdown AST并构建Word文档
            processNode(document, builder);
            
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(outputPath.toFile())) {
                wordDoc.write(out);
            }
            
            logger.debug("Word文档已保存: {}", outputPath);
            
        } catch (IOException e) {
            throw new ConversionException("创建Word文档失败", e);
        }
    }
    
    /**
     * 转换为PowerPoint演示文稿
     * 
     * @param document Markdown AST文档
     * @param outputPath 输出路径
     * @throws ConversionException 转换失败
     */
    private void convertToPowerPoint(Node document, Path outputPath) throws ConversionException {
        try (XMLSlideShow pptDoc = new XMLSlideShow()) {
            
            PowerPointBuilder builder = new PowerPointBuilder(pptDoc, config);
            
            // 遍历Markdown AST并构建PowerPoint文档
            processNode(document, builder);
            
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(outputPath.toFile())) {
                pptDoc.write(out);
            }
            
            logger.debug("PowerPoint文档已保存: {}", outputPath);
            
        } catch (IOException e) {
            throw new ConversionException("创建PowerPoint文档失败", e);
        }
    }
    
    /**
     * 处理Markdown AST节点
     * 
     * @param node AST节点
     * @param builder 文档构建器
     */
    private void processNode(Node node, DocumentBuilder builder) {
        for (Node child : node.getChildren()) {
            processNodeRecursive(child, builder);
        }
    }
    
    /**
     * 递归处理AST节点
     * 
     * @param node AST节点
     * @param builder 文档构建器
     */
    private void processNodeRecursive(Node node, DocumentBuilder builder) {
        if (node instanceof Heading) {
            Heading heading = (Heading) node;
            builder.addHeading(heading.getLevel(), getTextContent(heading));
        } else if (node instanceof Paragraph) {
            builder.addParagraph(getTextContent(node));
        } else if (node instanceof BulletList) {
            builder.startBulletList();
            for (Node listItem : node.getChildren()) {
                if (listItem instanceof BulletListItem) {
                    builder.addListItem(getTextContent(listItem));
                }
            }
            builder.endList();
        } else if (node instanceof OrderedList) {
            builder.startOrderedList();
            for (Node listItem : node.getChildren()) {
                if (listItem instanceof OrderedListItem) {
                    builder.addListItem(getTextContent(listItem));
                }
            }
            builder.endList();
        } else if (node instanceof FencedCodeBlock) {
            FencedCodeBlock codeBlock = (FencedCodeBlock) node;
            builder.addCodeBlock(codeBlock.getContentChars().toString());
        } else if (node instanceof BlockQuote) {
            builder.addBlockQuote(getTextContent(node));
        } else {
            // 递归处理子节点
            for (Node child : node.getChildren()) {
                processNodeRecursive(child, builder);
            }
        }
    }
    
    /**
     * 获取节点的文本内容
     * 
     * @param node AST节点
     * @return 文本内容
     */
    private String getTextContent(Node node) {
        StringBuilder text = new StringBuilder();
        extractText(node, text);
        return text.toString().trim();
    }
    
    /**
     * 递归提取文本内容
     * 
     * @param node AST节点
     * @param text 文本构建器
     */
    private void extractText(Node node, StringBuilder text) {
        if (node instanceof Text) {
            text.append(((Text) node).getChars());
        } else if (node instanceof SoftLineBreak || node instanceof HardLineBreak) {
            text.append(" ");
        } else {
            for (Node child : node.getChildren()) {
                extractText(child, text);
            }
        }
    }
    
    /**
     * 文档构建器接口
     */
    private interface DocumentBuilder {
        void addHeading(int level, String text);
        void addParagraph(String text);
        void startBulletList();
        void startOrderedList();
        void addListItem(String text);
        void endList();
        void addCodeBlock(String code);
        void addBlockQuote(String text);
    }
    
    /**
     * Word文档构建器
     */
    private static class WordDocumentBuilder implements DocumentBuilder {
        private final XWPFDocument document;
        private final Configuration config;
        
        public WordDocumentBuilder(XWPFDocument document, Configuration config) {
            this.document = document;
            this.config = config;
        }
        
        @Override
        public void addHeading(int level, String text) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(text);
            run.setBold(true);
            
            // 根据标题级别设置字体大小
            int fontSize = Math.max(12, 20 - level * 2);
            run.setFontSize(fontSize);
            
            paragraph.setSpacingAfter(200);
        }
        
        @Override
        public void addParagraph(String text) {
            if (!text.trim().isEmpty()) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(text);
                paragraph.setSpacingAfter(100);
            }
        }
        
        @Override
        public void startBulletList() {
            // Word中的列表处理在addListItem中实现
        }
        
        @Override
        public void startOrderedList() {
            // Word中的列表处理在addListItem中实现
        }
        
        @Override
        public void addListItem(String text) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("• " + text); // 简单的项目符号
            paragraph.setIndentationLeft(Units.toEMU(0.5)); // 缩进
        }
        
        @Override
        public void endList() {
            // 添加空段落作为分隔
            document.createParagraph();
        }
        
        @Override
        public void addCodeBlock(String code) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(code);
            run.setFontFamily("Consolas");
            run.setFontSize(10);
            paragraph.setStyle("Code");
            paragraph.setSpacingAfter(200);
        }
        
        @Override
        public void addBlockQuote(String text) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(text);
            run.setItalic(true);
            paragraph.setIndentationLeft(Units.toEMU(0.5));
            paragraph.setSpacingAfter(100);
        }
    }
    
    /**
     * PowerPoint构建器
     */
    private static class PowerPointBuilder implements DocumentBuilder {
        private final XMLSlideShow presentation;
        private final Configuration config;
        private XSLFSlide currentSlide;
        private XSLFTextBox currentTextBox;
        
        public PowerPointBuilder(XMLSlideShow presentation, Configuration config) {
            this.presentation = presentation;
            this.config = config;
        }
        
        @Override
        public void addHeading(int level, String text) {
            if (level == 1) {
                // 创建新幻灯片
                currentSlide = presentation.createSlide();
                XSLFTextBox titleBox = currentSlide.createTextBox();
                titleBox.setAnchor(new java.awt.Rectangle(50, 50, 600, 100));
                
                XSLFTextParagraph titleParagraph = titleBox.addNewTextParagraph();
                XSLFTextRun titleRun = titleParagraph.addNewTextRun();
                titleRun.setText(text);
                titleRun.setFontSize(24.0);
                titleRun.setBold(true);
                
                // 创建内容文本框
                currentTextBox = currentSlide.createTextBox();
                currentTextBox.setAnchor(new java.awt.Rectangle(50, 150, 600, 400));
            } else {
                // 子标题
                if (currentTextBox != null) {
                    XSLFTextParagraph paragraph = currentTextBox.addNewTextParagraph();
                    XSLFTextRun run = paragraph.addNewTextRun();
                    run.setText(text);
                    run.setFontSize(18.0);
                    run.setBold(true);
                }
            }
        }
        
        @Override
        public void addParagraph(String text) {
            if (currentTextBox != null && !text.trim().isEmpty()) {
                XSLFTextParagraph paragraph = currentTextBox.addNewTextParagraph();
                XSLFTextRun run = paragraph.addNewTextRun();
                run.setText(text);
                run.setFontSize(14.0);
            }
        }
        
        @Override
        public void startBulletList() {
            // PowerPoint中的列表处理在addListItem中实现
        }
        
        @Override
        public void startOrderedList() {
            // PowerPoint中的列表处理在addListItem中实现
        }
        
        @Override
        public void addListItem(String text) {
            if (currentTextBox != null) {
                XSLFTextParagraph paragraph = currentTextBox.addNewTextParagraph();
                XSLFTextRun run = paragraph.addNewTextRun();
                run.setText("• " + text);
                run.setFontSize(12.0);
                paragraph.setIndent(20.0);
            }
        }
        
        @Override
        public void endList() {
            // 添加空行
            if (currentTextBox != null) {
                currentTextBox.addNewTextParagraph();
            }
        }
        
        @Override
        public void addCodeBlock(String code) {
            if (currentTextBox != null) {
                XSLFTextParagraph paragraph = currentTextBox.addNewTextParagraph();
                XSLFTextRun run = paragraph.addNewTextRun();
                run.setText(code);
                run.setFontFamily("Consolas");
                run.setFontSize(10.0);
            }
        }
        
        @Override
        public void addBlockQuote(String text) {
            if (currentTextBox != null) {
                XSLFTextParagraph paragraph = currentTextBox.addNewTextParagraph();
                XSLFTextRun run = paragraph.addNewTextRun();
                run.setText(text);
                run.setFontSize(12.0);
                run.setItalic(true);
                paragraph.setIndent(20.0);
            }
        }
    }
}