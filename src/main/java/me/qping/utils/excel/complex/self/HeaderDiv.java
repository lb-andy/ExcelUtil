package me.qping.utils.excel.complex.self;

import lombok.Data;

import java.util.*;

/**
 * @ClassName HeaderDiv
 * @Description 矩形
 * @Author qping
 * @Date 2019/6/25 10:17
 * @Version 1.0
 **/
@Data
public class HeaderDiv {

    public static final int DIRECTION_RIGHT = 10;
    public static final int DIRECTION_BOTTOM = 20;

    // 位置大小
    int width = 0;
    int height = 0;

    // 值设定
    List<Cell> cellList = new ArrayList<>();
    List<Merge> mergeList = new ArrayList<>();

    Map<Integer, Integer> colWidthMap = new HashMap<>();


    public static HeaderDiv create(int width, int height){
        HeaderDiv div = new HeaderDiv();
        div.setWidth(width);
        div.setHeight(height);
        return div;
    }

    // 将整个div中的单元格全部合并
    public HeaderDiv merge(String value){
        return merge(value, null);
    }

    public HeaderDiv merge(String value, Style style){
        cellList.clear();
        mergeList.clear();
        return merge(0, 0, height - 1, width - 1, value, style);
    }

    public HeaderDiv merge(int beginRow, int beginCol, int endRow, int endCol, String value){
        return merge(beginRow, beginCol, endRow, endCol, value, null);
    }

    public HeaderDiv merge(int beginRow, int beginCol, int endRow, int endCol, String value, Style style){
        if(beginRow < 0 || beginRow > height - 1 || beginCol < 0 || beginCol > width - 1){
            throw new RuntimeException("begin merge 数组下标越界" + notNull(value));
        }
        if(endRow < 0 || endRow > height - 1 || endCol < 0 || endCol > width - 1){
            throw new RuntimeException("end merge 数组下标越界" + notNull(value));
        }
        Merge merge = Merge.create().begin(beginRow, beginCol).end(endRow, endCol).value(value);
        if(style != null){
            merge.setStyle(style);
            setMergeWidth(merge, style.getWidth());
        }

        mergeList.add(merge);
        return this;
    }

    private void setMergeWidth(Merge merge, int width){
        if(width <= 0) return;

        int beginCol = merge.getBegin().getCol();
        int endCol = merge.getEnd().getCol();

        // 将合并单元格的宽度拆分为 width / colCount
        int count = endCol - beginCol + 1;
        int eachWidth = width / count;
        for(int i = beginCol; i <= endCol; i++){
            colWidthMap.put(i, eachWidth);
        }
    }

    public HeaderDiv cell(int row, int col, String value){
        return cell(row, col, value, null);
    }

    public HeaderDiv cell(int row, int col, String value, Style style){
        if(row < 0 || row > height - 1 || col < 0 || col > width - 1){
            throw new RuntimeException("cell 数组下标越界" + notNull(value));
        }

        Cell cell = Cell.create().position(row, col).value(value);
        if(style != null){
            cell.setStyle(style);
        }
        if(style != null && style.getWidth() > 0){
            colWidthMap.put(col, style.getWidth());
        }
        cellList.add(cell);
        return this;
    }

    private static String notNull(Object o){
        if(o == null) return "";
        return o.toString();
    }


    public HeaderDiv append(HeaderDiv second, int direction) {

        if(direction == DIRECTION_RIGHT){

            int firstWidth = getWidth();

            // 将第二个放在第一个右边
            int width = firstWidth + second.getWidth();
            int height = getHeight() > second.getHeight() ? getHeight() : second.getHeight();

            setWidth(width);
            setHeight(height);

            for(Cell cell : second.getCellList()){
                // 将第二个的cell的col值往右移动
                Position pos = cell.getPosition();
                pos.setCol(pos.getCol() + firstWidth);
                this.getCellList().add(cell);
            }

            for(Merge merge : second.getMergeList()){
                Position begin = merge.getBegin();
                Position end = merge.getEnd();
                begin.setCol(begin.getCol() + firstWidth);
                end.setCol(end.getCol() + firstWidth);
                this.getMergeList().add(merge);
            }


            Iterator<Integer> keyItor = second.getColWidthMap().keySet().iterator();
            while(keyItor.hasNext()){
                Integer key = keyItor.next();
                colWidthMap.put(key + firstWidth, second.getColWidthMap().get(key));
            }


        }

        if(direction == DIRECTION_BOTTOM){

            int firstHeight = getHeight();

            // 将第二个放在第一个下边
            int height = firstHeight + second.getHeight();
            int width = getWidth() > second.getWidth() ? getWidth() : second.getWidth();

            setWidth(width);
            setHeight(height);

            for(Cell cell : second.getCellList()){
                // 将第二个的cell的row值往下移动
                Position pos = cell.getPosition();
                pos.setRow(pos.getRow() + firstHeight);
                this.getCellList().add(cell);
            }

            for(Merge merge : second.getMergeList()){
                Position begin = merge.getBegin();
                Position end = merge.getEnd();
                begin.setRow(begin.getRow() + firstHeight);
                end.setRow(end.getRow() + firstHeight);
                this.getMergeList().add(merge);
            }

            Iterator<Integer> keyItor = second.getColWidthMap().keySet().iterator();
            while(keyItor.hasNext()){
                Integer key = keyItor.next();
                colWidthMap.put(key, second.getColWidthMap().get(key));
            }
        }

        return this;

    }
}
