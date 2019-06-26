package me.qping.utils;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;

/**
 * @ClassName Ref
 * @Description TODO
 * @Author qping
 * @Date 2019/5/15 02:19
 * @Version 1.0
 **/
@Data
public class Ref {

    @Excel(name="药品名")
    String drug_name;

    @Excel(name ="平台药品名")
    String ref_name;

    @Excel(name="是否可以使用")
    int can_use;

}
