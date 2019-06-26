package me.qping.utils.excel.complex;

import lombok.extern.slf4j.Slf4j;
import me.qping.utils.excel.common.Config;
import me.qping.utils.excel.complex.self.*;
import me.qping.utils.excel.complex.self.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @ClassName ComplexUtil
 * @Description 复杂表头处理类
 * @Author qping
 * @Date 2019/6/25 11:16
 * @Version 1.0
 **/
@Slf4j
public class ComplexUtil {

    // 水平组合，从左往右依次布局
    public static HeaderDiv horizontal(HeaderDiv... headerDivs){
        if(headerDivs == null) return null;

        HeaderDiv first = headerDivs[0];
        for(int i = 1; i < headerDivs.length; i++){
            HeaderDiv second = headerDivs[i];
            first.append(second, HeaderDiv.DIRECTION_RIGHT);
        }
        return first;
    }

    public static HeaderDiv horizontal(List<HeaderDiv> headerDivs){
        if(headerDivs == null) return null;
        HeaderDiv[] array = new HeaderDiv[headerDivs.size()];
        return horizontal(headerDivs.toArray(array));
    }

    // 垂直组合
    public static HeaderDiv vertical(HeaderDiv... headerDivs){
        if(headerDivs == null) return null;

        HeaderDiv first = headerDivs[0];
        for(int i = 1; i < headerDivs.length; i++){
            HeaderDiv second = headerDivs[i];
            first.append(second, HeaderDiv.DIRECTION_BOTTOM);
        }
        return first;
    }

    public static HeaderDiv vertical(List<HeaderDiv> headerDivs){
        if(headerDivs == null) return null;
        HeaderDiv[] array = new HeaderDiv[headerDivs.size()];
        return vertical(headerDivs.toArray(array));
    }


    public static void draw(OutputStream outputStream, HeaderDiv complexHeader){

        Config config = new Config();
        config.initWorkbook("xls");

        Workbook workbook = config.getWorkbook();
        Sheet sheet = workbook.createSheet();


        // 考虑到复杂表头一般不会占用太大的内容，所以每一个 row 都初始化，便于操作
        for(int i = 0; i < complexHeader.getHeight(); i++ ){
            Row row = sheet.createRow(i);
        }

        for(me.qping.utils.excel.complex.self.Cell cell : complexHeader.getCellList()){
            org.apache.poi.ss.usermodel.Cell poiCell = sheet.getRow(cell.getPosition().getRow())
                    .createCell(cell.getPosition().getCol());
            poiCell.setCellValue(cell.getValue());

            if(cell.getStyle() != null){
                poiCell.setCellStyle(cell.getStyle().toCellStyle(workbook));
            }
        }




        for(Merge merge: complexHeader.getMergeList()){

            org.apache.poi.ss.usermodel.Cell poiCell = sheet.getRow(merge.getBegin().getRow())
                    .createCell(merge.getBegin().getCol());

            poiCell.setCellValue(merge.getValue());

            CellRangeAddress cellRange = new CellRangeAddress(
                    merge.getBegin().getRow(),
                    merge.getEnd().getRow(),
                    merge.getBegin().getCol(),
                    merge.getEnd().getCol()
            );

            sheet.addMergedRegion(cellRange);

            if(merge.getStyle() != null){
                poiCell.setCellStyle(merge.getStyle().toCellStyle(workbook));
                if(merge.getStyle().isBorder()){
                    RegionUtil.setBorderBottom(BorderStyle.THIN, cellRange, sheet); // 下边框
                    RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet); // 左边框
                    RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet); // 有边框
                    RegionUtil.setBorderTop(BorderStyle.THIN, cellRange, sheet); // 上边框
                }
            }
        }

        if(config.isAutoColumn()){
            for(int i = 0; i < complexHeader.getWidth(); i++){
                sheet.autoSizeColumn(i);
            }
        }else{
            Iterator<Integer> keyItor = complexHeader.getColWidthMap().keySet().iterator();
            while(keyItor.hasNext()){
                Integer col = keyItor.next();
                int width = 256 * complexHeader.getColWidthMap().get(col) + 184;
                sheet.setColumnWidth(col, width);
            }
        }

        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            log.error("export excel error");
        }
    }

//    public static void main(String[] args) {
//
//
//        Style titleStyle = StyleFactory.FONTBLOD_CENTER_WRAP_BORDER.copy().fontSize(18).fontFamily("宋体");
//        Style comStyle = StyleFactory.CENTER_WRAP_BORDER.copy().fontSize(12).fontFamily("宋体").width(5);
//        Style bigStyle = StyleFactory.FONTBLOD_CENTER_WRAP_BORDER.copy().fontSize(12).fontFamily("宋体").width(35);
//
//        String[] measureNames = new String[]{
//          "党建工作","妇幼卫生","公共卫生服务","健康管理","流动人口均等化服务等化服务等化服务等化服务等化服务等化服务等化服务等化服务等化服务", "满意度评价","内部管理","社区卫生","信息化建设","信用等级评价","合计"
//        };
//
//        List<HeaderDiv> secondDivs = new ArrayList<>();
//
//        HeaderDiv orgDiv = HeaderDiv.create(1, 3).merge("医疗单位", bigStyle);
//        secondDivs.add(orgDiv);
//
//        for(String name : measureNames){
//            HeaderDiv measure = HeaderDiv.create(2,3)
//                .merge(0, 0, 1,1, name, bigStyle)
//                .cell(2, 0, "应得", comStyle)
//                .cell(2, 1, "实得", comStyle);
//
//            secondDivs.add(measure);
//        }
//
//        HeaderDiv secondDiv = ComplexUtil.horizontal(secondDivs);
//
//        HeaderDiv title = HeaderDiv.create(secondDiv.getWidth(), 2)
//                .merge("关于开展2018年第四季度综合目标完成情况考核的通知医疗机构成绩汇总", titleStyle);
//
//
//
//        HeaderDiv complexHeader = ComplexUtil.vertical(title, secondDiv);
//
//        try {
//            ComplexUtil.draw(new FileOutputStream(new File("/Users/qping/Desktop/1.xls")), complexHeader);
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//
//
//    }

}
