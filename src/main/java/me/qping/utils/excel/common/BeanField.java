package me.qping.utils.excel.common;

import lombok.Data;
import me.qping.utils.excel.anno.Excel;

import java.lang.reflect.Field;

/**
 * @ClassName BeanField
 * @Author qping
 * @Date 2019/5/10 09:19
 * @Version 1.0
 **/
@Data
public class BeanField implements Comparable<BeanField>{

    Field field;
    String name = null;
    String dateformat = null;
    int index = -1;
    int sort = -1;

    boolean userDefineIndex = false;

    public BeanField(Field field, Excel excel) {
        this.index = excel.index();
        this.name = excel.name();
        this.sort = excel.sort();
        this.field = field;

        if(this.index > -1){
            this.userDefineIndex = true;
        }
    }

    @Override
    public int compareTo(BeanField o) {
        return this.sort == o.sort ? 0 :
                this.sort > o.sort ? 1 : -1;
    }
}
