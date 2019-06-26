package me.qping.utils.csv;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;

import java.io.*;
import java.util.Arrays;
import java.util.Date;

/**
 * @ClassName CsvUtil
 * @Description TODO
 * @Author qping
 * @Date 2019/6/26 09:00
 * @Version 1.0
 **/
public class CsvUtil {

    public static void main(String[] args) throws IOException {


        String pathString = "/Users/qping/Downloads/儿童医院、鼓楼医院2019年1-5月数据/";

        File path = new File(pathString);

        File[] files = path.listFiles();

        JSONArray jsonArray = new JSONArray();
        for(File file : files){

            String fileName = file.getName();
            InputStream csvInputStream = new FileInputStream( file);
            BufferedReader reader = new BufferedReader(new InputStreamReader(csvInputStream, "GBK"));
            Iterable<CSVRecord> records = CSVFormat.DEFAULT.withFirstRecordAsHeader().parse(reader);

            int i = 1;
            for (CSVRecord record : records) {
                String 病史描述 = record.get("病史描述");
                String 主诉 = record.get("主诉");
                String 身份证号 = record.get("身份证号");

                身份证号 = 身份证号.startsWith("'") ? 身份证号.substring(1, 身份证号.length()) : 身份证号;

                JSONObject row = new JSONObject();
                row.put("vid", fileName + "_" + i + "_" + 身份证号);
                row.put("zs", 主诉);
                row.put("bs", 病史描述);

                jsonArray.add(row);

                i++;
            }
        }

        System.out.println(jsonArray.size());

        FileWriter writer  = new FileWriter("/Users/qping/Downloads/1.json" );
        writer.write(jsonArray.toJSONString());
        writer.flush();
        writer.close();

    }
}
