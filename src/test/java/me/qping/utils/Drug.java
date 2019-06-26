package me.qping.utils;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;

/**
 * @ClassName Drug
 * @Description TODO
 * @Author qping
 * @Date 2019/5/10 10:40
 * @Version 1.0
 **/
@Data
public class Drug {

    @Excel(name = "药物目录", sort = 1)
    String generalName;

    @Excel(name="商品名")
    String drugName;


    @Excel(name = "平台药物名称", sort = 2)
    String drugPlatName;

    @Excel(index= 2)
    String spec;



}