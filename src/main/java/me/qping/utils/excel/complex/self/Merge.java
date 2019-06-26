package me.qping.utils.excel.complex.self;

import lombok.Data;

/**
 * @ClassName Merge
 * @Description 合并单元格
 * @Author qping
 * @Date 2019/6/25 10:57
 * @Version 1.0
 **/
@Data
public class Merge {
    Position begin = new Position();
    Position end = new Position();
    String value;
    Style style;

    public static Merge create(){
        return new Merge();
    }

    public Merge begin(int row, int col){
        this.begin.setRow(row);
        this.begin.setCol(col);
        return this;
    }

    public Merge end(int row, int col){
        this.end.setRow(row);
        this.end.setCol(col);
        return this;
    }

    public Merge value(String value){
        this.value = value;
        return this;
    }

    public Merge style(Style style){
        this.style = style;
        return this;
    }


}
