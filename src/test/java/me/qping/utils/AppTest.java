package me.qping.utils;

import static org.junit.Assert.assertTrue;

import me.qping.utils.excel.ExcelUtil;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.common.Consumer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.Collection;
import java.util.List;

/**
 * Unit test for simple App.
 */
public class AppTest {

    @Test
    public void importExcel() throws FileNotFoundException {
        ExcelUtil excelUtil = new ExcelUtil();
//        FileInputStream inputStream = new FileInputStream("/Users/qping/Documents/监室问诊/药物目录.xls");
//        Collection<Drug> data = excelUtil.read(Drug.class, inputStream);
//        Config context = excelUtil.getConfig();
//        saveToDB(data);

        FileInputStream inputStream = new FileInputStream("/Users/qping/Documents/监室问诊/疾病药品对照(1).xls");
        excelUtil.readEachSheet(Ref.class, inputStream, new Consumer<Ref>() {
            @Override
            public void execute(Collection<Ref> data, Config context) {
//                saveRelationToDB(data, context.getHeaders().get(3), context.getHeaders().get(4));
                System.out.println(data.size());
            }
        });

    }

    private <T> void saveRelationToDB(Collection<Ref> data, String diseaseCode, String diseaseName) {
        try {
            Class.forName("com.mysql.jdbc.Driver");
        } catch (ClassNotFoundException e) {
            return;
        }

        String dbURL = "jdbc:mysql://192.168.100.17:3306/datamiller?useUnicode=true&characterEncoding=UTF-8&tinyInt1isBit=false";
        String userName = "datamiller";
        String userPwd = "datamiller";

        String sql = "insert into t_prison_drug_use(drug_id, disease_code, disease_name) values(%d,'%s','%s')";
        try (
                Connection dbConn = DriverManager.getConnection(dbURL, userName, userPwd);
        ){
            dbConn.setAutoCommit(false);
            for(int i = 0; i < data.size(); i++){
                Ref ref = ((List<Ref>)data).get(i);

                if(ref.can_use != 1 || i <= 4){
                    continue;
                }
                String t_sql = String.format(sql, i + 1,
                        diseaseCode,
                        diseaseName);

                PreparedStatement ps = dbConn.prepareStatement(t_sql);
                ps.execute();
            }
            dbConn.commit();
//
//            ps.execute();
//            ResultSet rs = ps.getResultSet();

            System.out.println("连接数据库成功");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.print("连接失败");
        }


        System.out.println(diseaseName);
    }

    private void saveToDB(Collection<Drug> data) {
        try {
            Class.forName("com.mysql.jdbc.Driver");
        } catch (ClassNotFoundException e) {
            return;
        }

        String dbURL = "jdbc:mysql://192.168.100.17:3306/datamiller?useUnicode=true&characterEncoding=UTF-8&tinyInt1isBit=false";
        String userName = "datamiller";
        String userPwd = "datamiller";

        String sql = "insert into t_prison_drug_dict(general_name, drug_name, spec) values('%s','%s','%s')";
        try (
            Connection dbConn = DriverManager.getConnection(dbURL, userName, userPwd);
        ){
            dbConn.setAutoCommit(false);
            for(Drug drug : data){
                String t_sql = String.format(sql, drug.getGeneralName(),
                        drug.getDrugName() == null ? "" : drug.getDrugName(),
                        drug.getSpec() == null ? "" : drug.getSpec());

                PreparedStatement ps = dbConn.prepareStatement(t_sql);
                ps.execute();
            }
            dbConn.commit();
//
//            ps.execute();
//            ResultSet rs = ps.getResultSet();

            System.out.println("连接数据库成功");
        } catch (Exception e) {
            e.printStackTrace();
            System.out.print("连接失败");
        }
    }

    @Test
    public void shouldAnswerWithTrue() throws IOException {
        ExcelUtil excelUtil = new ExcelUtil();
        Collection<Disease> diseases = excelUtil.read(Disease.class, "/Users/qping/Documents/监室问诊/三百多种疾病支持列表.xlsx");
        Collection<Drug> drugs = excelUtil.read(Drug.class, "/Users/qping/Documents/监室问诊/药物目录.xls");

        System.out.println(diseases.size());

        OutputStream outputStream = new FileOutputStream("/Users/qping/Documents/监室问诊/疾病药品对照.xls");

        Workbook workbook =  new HSSFWorkbook();

        for(Disease disease : diseases){

            Sheet sheet = workbook.createSheet(disease.getIcd10().replace("*", ""));

            Row title = sheet.createRow(0);
            title.createCell(0).setCellValue("药品名");
            title.createCell(1).setCellValue("平台药品名");
            title.createCell(2).setCellValue("是否可以使用");
            title.createCell(3).setCellValue(disease.getIcd10());
            title.createCell(4).setCellValue(disease.getIcdName());

            int rowIndex = 0;
            for(Drug drug: drugs){
                rowIndex ++;
                Row row = sheet.createRow(rowIndex);
                row.createCell(0).setCellValue(drug.getDrugName());
                row.createCell(1).setCellValue(drug.getDrugPlatName());
            }

            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
        }

        workbook.write(outputStream);
    }


}
