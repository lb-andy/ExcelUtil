package me.qping.utils.excel.complex.self;

import lombok.Data;

/**
 * @ClassName Cell
 * @Description 单元格
 * @Author qping
 * @Date 2019/6/25 11:10
 * @Version 1.0
 **/
@Data
public class Cell {

    Position position = new Position();
    String value;
    Style style;

    public static Cell create(){
        return new Cell();
    }

    public Cell position(int row, int col){
        this.position.setRow(row);
        this.position.setCol(col);
        return this;
    }

    public Cell position(Position begin){
        this.position = begin;
        return this;
    }

    public Cell value(String value){
        this.value = value;
        return this;
    }

    public Cell style(Style style){
        this.style = style;
        return this;
    }



}
