package me.xueyao.utils;


import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.*;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.style.RtfParagraphStyle;
import lombok.Getter;
import lombok.Setter;

import java.awt.*;
import java.io.FileOutputStream;

/**
 * @author Simon.Xue
 * @date 2019-12-04 21:47
 **/
@Getter
@Setter
public class WordUtils {
    private Document document;
    private BaseFont bfChinese;

    public WordUtils() {
        document = new Document(PageSize.A4);
    }

    /**
     * 创建一个Writer与document对象关联
     * @param filePath 文档路径，没有则创建
     * @throws Exception
     */
    public void openDocument(String filePath) throws Exception{
        RtfWriter2.getInstance(this.document,
                new FileOutputStream(filePath));
        this.document.open();
        this.bfChinese = BaseFont.createFont("Times-Roman",
                "winansi", BaseFont.NOT_EMBEDDED);
    }

    /**
     * 添加标题
     * @param titleStr  标题
     * @param fontSize  字体大小
     * @param fontStyle  字体样式
     * @param elementAlign  对齐方式
     * @throws Exception
     */
    public void insertTitle(String titleStr, int fontSize,
                            int fontStyle, int elementAlign) throws Exception{
        Font titleFont = new Font((this.bfChinese), fontSize, fontStyle);
        Paragraph title = new Paragraph(titleStr);
        //设置标题格式对齐方式
        title.setAlignment(elementAlign);
        title.setFont(titleFont);
        this.document.add(title);
    }

    /**
     * 设置带有目录格式的标题(标题1格式)
     * @param titleStr
     * @param rtfParagraphStyle
     * @throws Exception
     */
    public void insertTitlePattern(String titleStr,
                                   RtfParagraphStyle rtfParagraphStyle) throws Exception{
        Paragraph title = new Paragraph(titleStr);
        title.setFont(rtfParagraphStyle);
        this.document.add(title);
    }

    /**
     * 设置带目录格式的标题(标题2格式)
     * @param titleStr
     * @param rtfParagraphStyl
     * @throws DocumentException
     */
    public void insertTitlePatternSecond(String titleStr,
                                         RtfParagraphStyle rtfParagraphStyl) throws DocumentException {
        Paragraph title = new Paragraph(titleStr);
        //设置标题格式对齐方式
        title.setFont(rtfParagraphStyl);
        this.document.add(title);
    }

    /**
     * 添加表名
     * @param tableName
     * @param fontSize
     * @param fontStyle
     * @param elementAlign
     * @throws DocumentException
     */
    public void insertTableName(String tableName, int fontSize,
                                int fontStyle, int elementAlign) throws DocumentException {
        Font titleFont = new Font(this.bfChinese, fontSize, fontStyle);
        titleFont.setColor(102, 102, 153);
        Paragraph title = new Paragraph(tableName);
        //设置标题格式对齐方式
        title.setAlignment(elementAlign);
        title.setFont(titleFont);

        this.document.add(title);
    }

    /**
     * 添加内容
     * @param contextStr  内容部分
     * @param fontSize  字体大小
     * @param fontStyle  字体样式
     * @param elementAlign  对齐方式
     * @throws DocumentException
     */
    public void insertContext(String contextStr, int fontSize,
                              int fontStyle, int elementAlign) throws DocumentException {
        Font contentFont = new Font(bfChinese, fontSize, fontStyle);
        Paragraph context = new Paragraph(contextStr);
        //设置行距
        context.setLeading(3f);
        //正文格式左对齐
        //context.setAlignment(elementAlign);
        context.setFont(contentFont);
        //离上一段落(标题)空的行数
        context.setSpacingBefore(1);
        //设置第一行空的列数
        context.setFirstLineIndent(20);
        this.document.add(context);
    }

    /**
     * 添加图片
     * @param imgUrl 图片路径
     * @param imageAlign  显示位置
     * @param height  显示高度
     * @param weight  显示宽度
     * @param percent  显示比例
     * @param heightPercent  显示高度比例
     * @param weightPercent  显示宽度比例
     * @param rotation  显示图片旋转角度
     * @throws Exception
     */
    public void insertImg(String imgUrl, int imageAlign, int height,
                          int weight, int percent, int heightPercent,
                          int weightPercent, int rotation) throws Exception {
        Image img = Image.getInstance(imgUrl);
        if (null == img) {
            return;
        }
        img.setAbsolutePosition(0, 0);
        img.setAlignment(imageAlign);
        img.scaleAbsolute(height, weight);
        img.scaleAbsolute(1000, 1000);
        img.scalePercent(percent);
        img.scalePercent(heightPercent, weightPercent);
        img.setRotation(rotation);
        document.add(img);
    }

    /**
     * 添加简单表格
     * @param column 表格列数(必须)
     * @param row  表格行数
     * @throws Exception
     */
    public void insertSimpleTable(int column, int row) throws Exception {
        //设置列数
        Table table = new Table(column);
        //居中显示
        table.setAlignment(Element.ALIGN_CENTER);
        //纵向居中显示
        table.setAlignment(Element.ALIGN_MIDDLE);
        //自动填满
        table.setAutoFillEmptyCells(true);
        //设置边框颜色
        table.setBorderColor(new Color(0, 125, 255));
        //边框宽度
        table.setBorderWidth(1);
        //衬距
        table.setSpacing(2);
        //单元格之间的间距
        table.setPadding(2);
        //边框
        table.setBorder(20);

        for (int i = 0; i < column * 3; i++) {
            table.addCell(new Cell("" + i));
        }
        document.add(table);
    }

    /**
     * 关闭document
     */
    public void closeDocument() {
        this.document.close();
        document = null;
    }
}
