package me.qping.utils.excel.complex.self;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;

/**
 * @ClassName Style
 * @Description 样式封装
 * @Author qping
 * @Date 2019/6/25 16:34
 * @Version 1.0
 **/
@Data
public class Style {

    Color backgroundColor;

    Boolean fontBold;
    Color fontColor;
    Short fontSize;
    String fontFamily;

    boolean center;
    boolean wrap;
    boolean border;

    int width;

    public Style backgroundColor(Color backgroundColor){
        this.backgroundColor = backgroundColor;
        return this;
    }

    public Style fontBold(Boolean fontBold){
        this.fontBold = fontBold;
        return this;
    }

    public Style fontColor(Color fontColor){
        this.fontColor = fontColor;
        return this;
    }

    public Style fontSize(Short fontSize){
        this.fontSize = fontSize;
        return this;
    }

    public Style fontSize(int fontSize){
        this.fontSize = (short) fontSize;
        return this;
    }

    public Style fontFamily(String fontFamily){
        this.fontFamily = fontFamily;
        return this;
    }

    public Style center(boolean center){
        this.center = center;
        return this;
    }

    public Style wrap(boolean wrap){
        this.wrap = wrap;
        return this;
    }

    public Style border(boolean border){
        this.border = border;
        return this;
    }

    public Style width(int width){
        this.width = width;
        return this;
    }

    public Style copy(){
        Style style = new Style();
        style.setBackgroundColor(backgroundColor);
        style.setFontBold(fontBold);
        style.setFontColor(fontColor);
        style.setFontSize(fontSize);
        style.setFontFamily(fontFamily);
        style.setCenter(center);
        style.setWrap(wrap);
        style.setBorder(border);
        return style;
    }

    public CellStyle toCellStyle(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();

        if(backgroundColor != null){
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if(workbook instanceof HSSFWorkbook){
                HSSFPalette palette = ((HSSFWorkbook)workbook).getCustomPalette();
                HSSFColor myColor = palette.findSimilarColor(backgroundColor.getRed(), backgroundColor.getGreen(), backgroundColor.getBlue());
                ((HSSFCellStyle)cellStyle).setFillForegroundColor(myColor.getIndex());
            }
            if(workbook instanceof XSSFWorkbook){
                XSSFColor myColor = new XSSFColor(backgroundColor);
                ((XSSFCellStyle) cellStyle).setFillForegroundColor(myColor);
            }
        }


        if(fontColor != null || fontSize != null || fontFamily != null || fontBold != null){
            Font fontStyle = workbook.createFont();
            if(fontBold != null){
                fontStyle.setBold(fontBold);
            }
            if(fontColor != null){
                if(workbook instanceof HSSFWorkbook){
                    HSSFPalette palette = ((HSSFWorkbook)workbook).getCustomPalette();
                    HSSFColor myColor = palette.findSimilarColor(backgroundColor.getRed(), backgroundColor.getGreen(), backgroundColor.getBlue());
                    ((HSSFFont)fontStyle).setColor(myColor.getIndex());
                }
                if(workbook instanceof XSSFWorkbook){
                    XSSFColor myColor = new XSSFColor(backgroundColor);
                    ((XSSFFont) fontStyle).setColor(myColor);
                }
            }
            if(fontSize != null){
                fontStyle.setFontHeightInPoints(fontSize);
            }
            if(fontFamily != null){
                fontStyle.setFontName(fontFamily);
            }
            cellStyle.setFont(fontStyle);
        }

        if(center){
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
        }

        if(wrap){
            cellStyle.setWrapText(true);
        }

        // 列宽
        // sheet.setColumnWidth(0, 3766);

        if(border){
            cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
            cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
            cellStyle.setBorderTop(BorderStyle.THIN);//上边框
            cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        }

        return cellStyle;
    }

}
