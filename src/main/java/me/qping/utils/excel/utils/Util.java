package me.qping.utils.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * @ClassName Util
 * @Author qping
 * @Date 2019/5/15 03:18
 * @Version 1.0
 **/
public class Util {

    public static boolean isBlank(String str) {
        return str == null || "".equals(str);
    }

    public static boolean isNotBlank(String str) {
        return !isBlank(str);
    }

    public static Object getCellValue(Cell cell) {


        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();

        switch (cellType){
            case _NONE:
                return null;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                } catch (IllegalStateException e) {
                    return cell.getRichStringCellValue();
                }
            case BLANK:
                return null;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
        }
        return null;
    }


}
