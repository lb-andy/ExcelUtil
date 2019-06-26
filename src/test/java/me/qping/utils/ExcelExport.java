package me.qping.utils;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;

/**
 * @ClassName ExcelExport
 * @Description TODO
 * @Author qping
 * @Date 2019/5/13 16:54
 * @Version 1.0
 **/
@Data
public class ExcelExport {

    @Excel(name = "药物目录", sort = 1)
    String drugName;

    @Excel(name = "平台药物名称", sort = 2)
    String drugPlatName;

    @Excel(name = "是否可以使用", sort = 3)
    String use;

}
