package me.qping.utils.excel.complex.self;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;

/**
 * @ClassName Style工厂
 * @Description 样式复用
 * @Author qping
 * @Date 2019/6/25 16:34
 * @Version 1.0
 **/
@Data
public class StyleFactory {

    public static Style CENTER;
    public static Style CENTER_BORDER;
    public static Style CENTER_WRAP;
    public static Style CENTER_WRAP_BORDER;


    public static Style FONTBLOD_CENTER;
    public static Style FONTBLOD_CENTER_BORDER;
    public static Style FONTBLOD_CENTER_WRAP;
    public static Style FONTBLOD_CENTER_WRAP_BORDER;

    static{
        CENTER = new Style().center(true);
        CENTER_BORDER = CENTER.copy().border(true);
        CENTER_WRAP = CENTER.copy().wrap(true);
        CENTER_WRAP_BORDER = CENTER_WRAP.copy().border(true);

        FONTBLOD_CENTER = new Style().fontBold(true).center(true);
        FONTBLOD_CENTER_BORDER = FONTBLOD_CENTER.copy().border(true);
        FONTBLOD_CENTER_WRAP = FONTBLOD_CENTER.copy().wrap(true);
        FONTBLOD_CENTER_WRAP_BORDER = FONTBLOD_CENTER.copy().wrap(true).border(true);
    }



}
