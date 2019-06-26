package me.qping.utils.excel.common;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.anno.Excel;
import me.qping.utils.excel.utils.Util;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @ClassName Config
 * @Author qping
 * @Date 2019/5/15 03:05
 * @Version 1.0
 **/
@Data
@Slf4j
public class Config {

    public static final int ERROR_STRATE_BREAK = 1;     // stop when error
    public static final int ERROR_STRATE_CONTINUE = 2;  // continue when error

    // global config
    boolean firstHeader = true;
    int errorStrategy = ERROR_STRATE_CONTINUE;

    // runtime config
    int sheetNo = 0;
    String sheetName;
    List<BeanField> beanFields;
    List<String> headers;
    Workbook workbook;

    String fileExt;
    Class clazz;

    boolean autoColumn = false;


    public <T> void init(Class<T> clazz) {
        this.beanFields = getExcelFields(clazz);
        this.clazz = clazz;
    }

    public void initWorkbook(InputStream inputStream){
        try {
            workbook = WorkbookFactory.create(inputStream);
            if(workbook instanceof HSSFWorkbook){
                fileExt = "xls";
            }else{
                fileExt = "xlsx";
            }
        } catch (IOException e) {
            log.error("load excel error!!!");
        }
    }

    public void initWorkbook(String fileExt){
        if(fileExt.equals("xls")){
            workbook = new HSSFWorkbook();
        }else{
            workbook = new XSSFWorkbook();
        }
        this.fileExt = fileExt;
    }

    public void initHeader() {
        initHeader(sheetNo);
    }

    public <T> void initHeader(int sheetNo) {

        Sheet sheet = workbook.getSheetAt(sheetNo);
        sheetName = sheet.getSheetName();

        // read first row as title
        headers = new ArrayList<>();
        if(firstHeader){
            Row headerRow = sheet.getRow(0);
            int colNum = headerRow.getPhysicalNumberOfCells();
            for(int i = 0; i < colNum; i++){
                Cell headerCell = headerRow.getCell(i);
                Object headerTitle = Util.getCellValue(headerCell);
                if(headerTitle == null){
                    continue;
                }
                headers.add(headerTitle.toString());
            }
        }

        // transfer name to index
        for(BeanField beanField : beanFields){
            // if index and name is set in the same time, index first
            if(Util.isNotBlank(beanField.getName()) && !beanField.userDefineIndex){
                int index = headers.indexOf(beanField.getName());
                beanField.setIndex(index);
            }
        }
    }

    public static <T> List<BeanField> getExcelFields(Class<T> clazz){
        List<BeanField> beanFields = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        for(Field field : fields){
            Excel excel = field.getAnnotation(Excel.class);
            if(excel == null){
                continue;
            }
            field.setAccessible(true);
            BeanField beanField = new BeanField(field, excel);
            beanFields.add(beanField);
        }

        Collections.sort(beanFields);
        return beanFields;
    }


}
