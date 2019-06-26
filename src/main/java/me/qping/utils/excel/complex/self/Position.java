package me.qping.utils.excel.complex.self;

import lombok.Data;

/**
 * @ClassName Position
 * @Description 位置，记录单元格的 行和列，从0开始
 * @Author qping
 * @Date 2019/6/25 10:52
 * @Version 1.0
 **/
@Data
public class Position {
    int row;
    int col;

    public static Position of(int row, int col) {
        Position pos = new Position();
        pos.setRow(row);
        pos.setCol(col);
        return pos;
    }
}
