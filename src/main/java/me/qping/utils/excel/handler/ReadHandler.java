package me.qping.utils.excel.handler;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.BeanField;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.common.Consumer;
import me.qping.utils.excel.utils.Util;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

/**
 * @ClassName ReadHandler
 * @Author qping
 * @Date 2019/5/15 03:04
 * @Version 1.0
 **/
@Slf4j
public class ReadHandler {

    public void eachSheet(Config config, Consumer consumer) {
        int sheetCount = config.getWorkbook().getNumberOfSheets();
        for(int sheetno = 0; sheetno < sheetCount; sheetno++){

            config.setSheetNo(sheetno);
            config.initHeader(sheetno);

            List<T> list = new ArrayList<>();
            try {
                transferSheetData(config, config.getWorkbook().getSheetAt(sheetno), list, config.getClazz(), config.getBeanFields());
            } catch (Exception e) {
                e.printStackTrace();
            }
            consumer.execute(list, config);
        }

    }

    public <T> Collection<T> read(Config config) {
        config.initHeader();
        // 读取数据转换为bean
        List<T> list = new ArrayList<>();
        try {
            transferSheetData(config, config.getWorkbook().getSheetAt(config.getSheetNo()), list, config.getClazz(), config.getBeanFields());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    private <T> void transferSheetData(Config config, Sheet sheet, List<T> list, Class<T> clazz, List<BeanField> beanFields) throws Exception {
        Iterator<Row> rowIt = sheet.rowIterator();

        if(config.isFirstHeader()){
            if(rowIt.hasNext()){
                rowIt.next();
            }
        }

        while(rowIt.hasNext()){
            Row row = rowIt.next();
            T obj = clazz.newInstance();
            for(BeanField beanField : beanFields){
                int colIndex = beanField.getIndex();

                if(colIndex == -1){
                    continue;
                }

                Cell cell = row.getCell(colIndex);
                Object cellValue = Util.getCellValue(cell);

                setValue(obj, beanField, cellValue);

            }
            list.add(obj);
        }
    }

    private <T>  void setValue(T obj, BeanField beanField, Object cellValue) throws Exception {

        if(cellValue == null){
            return;
        }

        Field field = beanField.getField();
        Class<?> type = field.getType();
        try{

            // 类型为String
            if(type == String.class){
                // todo 日期格式化处理
                beanField.getField().set(obj, cellValue);
            }
            // 类型为Int
            else if(type == Integer.class || type == int.class ){
                if(cellValue instanceof String){
                    int val = Integer.parseInt((String)cellValue);
                    beanField.getField().set(obj, val);
                }
                else if(cellValue instanceof Double){
                    int val = ((Double) cellValue).intValue();
                    beanField.getField().set(obj, val);
                }else{
                    beanField.getField().set(obj, cellValue);
                }
            }else{
                beanField.getField().set(obj, cellValue);
            }

        }catch (Exception ex){
            throw new Exception("值转换错误，属性名称：" + field.getName()
                    + ", 期望类型：" + field.getType().toString()
                    + ", 实际类型：" + cellValue.getClass().getTypeName()
                    + ", 实际值为：" + cellValue
            );
        }


    }


}
