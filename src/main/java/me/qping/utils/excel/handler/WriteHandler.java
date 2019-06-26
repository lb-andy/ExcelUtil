package me.qping.utils.excel.handler;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.BeanField;
import me.qping.utils.excel.common.Config;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;

/**
 * @ClassName WriteHandler
 * @Author qping
 * @Date 2019/5/15 03:05
 * @Version 1.0
 **/
@Slf4j
public class WriteHandler {
    public <T> void write(Config config, OutputStream outputStream, Collection<T> data) {
        Sheet sheet = config.getWorkbook().createSheet();
        int rowIndex = -1;
        for(T rowData : data){
            rowIndex++;

            Row row = sheet.createRow(rowIndex);

            int colIndex = -1;


            for(BeanField beanField : config.getBeanFields()){
                if(beanField.getName() == null){
                    continue;
                }

                colIndex++;

                try {
                    Object valueObj = beanField.getField().get(rowData);
                    String value = valueObj == null ? "" : valueObj.toString();
                    Cell cell = row.createCell(colIndex);
                    cell.setCellValue(value);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }

            }
        }
        try {
            config.getWorkbook().write(outputStream);
        } catch (IOException e) {
            log.error("export excel error");
        }
    }

}
