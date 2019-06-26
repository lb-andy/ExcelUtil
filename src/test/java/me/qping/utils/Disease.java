package me.qping.utils;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;

/**
 * @ClassName Disease
 * @Description TODO
 * @Author qping
 * @Date 2019/5/13 16:46
 * @Version 1.0
 **/
@Data
public class Disease {

    @Excel(name = "疾病ICD-10")
    String icd10;

    @Excel(name = "疾病名称")
    String icdName;


}
