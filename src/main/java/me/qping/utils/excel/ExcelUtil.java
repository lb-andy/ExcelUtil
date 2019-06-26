package me.qping.utils.excel;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.Consumer;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.handler.ReadHandler;
import me.qping.utils.excel.handler.WriteHandler;

import java.io.*;
import java.util.*;

/**
 * @ClassName ExcelUtil
 * @Description excel读取封装
 * @Author qping
 * @Date 2019/5/9 17:11
 * @Version 1.0
 **/
@Slf4j
@Data
public class ExcelUtil {

    private Config config = new Config();

    // 抽离ExcelUtil业务代码，便于阅读
    private WriteHandler writeHandler = new WriteHandler();
    private ReadHandler readHandler = new ReadHandler();


    public <T> void write(Class<T> clazz, String filePath, Collection<T> data) throws FileNotFoundException {
        String fileExt = "xls";
        if(filePath.endsWith(".xlsx")){
            fileExt = "xlsx";
        }
        this.write(clazz, new FileOutputStream(filePath), data, fileExt);

        try {
            config.getWorkbook().close();
        } catch (IOException e) { }
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data) {
        this.write(clazz, outputStream, data, "xls");
    }

    public <T> void write(Class<T> clazz, OutputStream outputStream, Collection<T> data, String fileExt) {
        config.init(clazz);
        config.initWorkbook(fileExt);
        writeHandler.write(config, outputStream, data);
    }

    public <T> Collection<T> read(Class<T> clazz, String filePath) throws FileNotFoundException {
        Collection<T> list = this.read(clazz, new FileInputStream(new File(filePath)));

        try {
            config.getWorkbook().close();
        } catch (IOException e) { }

        return list;
    }

    public <T> Collection<T> read(Class<T> clazz, InputStream inputStream){
        config.init(clazz);
        config.initWorkbook(inputStream);
        return readHandler.read(config);
    }

    public <T> void readEachSheet(Class<T> clazz, InputStream inputStream, Consumer<T> consumer){
        config.init(clazz);
        config.initWorkbook(inputStream);
        readHandler.eachSheet(config, consumer);
    }

    public ExcelUtil firstHeader(boolean firstHeader){
        this.config.setFirstHeader(firstHeader);
        return this;
    }

    public ExcelUtil sheetNo(int sheetNo){
        this.config.setSheetNo(sheetNo);
        return this;
    }

}
