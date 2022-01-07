package ziv.excel.news.imports;

import cn.hutool.poi.excel.cell.CellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.HashMap;
import java.util.Map;

public class PoiCellParser {


    /**
     * 根据行获取表头
     *
     * @param row 表头行
     * @return 表头每一列对应的值和索引
     */
    public static Map<Integer, PoiCell> parseCellByRow(Row row) {
        Map<Integer, PoiCell> headerMapping = new HashMap<>();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            PoiCell cellValue = PoiCellParser.getCellValue(cell);
            headerMapping.put(i, cellValue);
        }
        return headerMapping;
    }

    /**
     * 获取CmsCell对象
     *
     * @param cell cell对象
     * @return 自定义对象
     */
    public static PoiCell getCellValue(Cell cell) {
        Object cellValue = CellUtil.getCellValue(cell, true);
        PoiCell poiCell = new PoiCell();
        CellType cellType = cell.getCellType();
        poiCell.setCellType(cellType.getCode());
        poiCell.setJavaType(cellValue == null ? null : cellValue.getClass());
        poiCell.setValue(cellValue);
        return poiCell;
    }

    /**
     * 将对象类型转换成目标类型看，如果可能的话
     *
     * @param object 被操作的对象
     * @param clazz  目标类型
     * @return Object
     */
    public static Object converter(Object object, Class<?> clazz) {
        if (clazz.equals(String.class)) {
            return object.toString();
        } else if (clazz.equals(Boolean.class) || boolean.class.equals(clazz)) {
            return Boolean.parseBoolean(object.toString());
        } else if (clazz.equals(Double.class) || double.class.equals(clazz)) {
            return Double.parseDouble(object.toString());
        } else if (clazz.equals(Float.class) || float.class.equals(clazz)) {
            return Float.parseFloat(object.toString());
        } else if (clazz.equals(Integer.class) || int.class.equals(clazz)) {
            return Integer.parseInt(object.toString());
        } else if (clazz.equals(Short.class) || short.class.equals(clazz)) {
            return Short.parseShort(object.toString());
        } else if (clazz.equals(Byte.class) || byte.class.equals(clazz)) {
            return Byte.parseByte(object.toString());
        } else if (clazz.equals(Character.class) || char.class.equals(clazz)) {
            return object.toString().charAt(0);
        }
        return object;
    }
}
