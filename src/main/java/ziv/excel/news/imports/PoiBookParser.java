package ziv.excel.news.imports;

import cn.hutool.core.lang.Assert;
import org.apache.commons.compress.utils.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import ziv.excel.news.PoiConfig;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

public class PoiBookParser {

    ThreadLocal<Map<String, Object>> errorThreadMap = new ThreadLocal<>();

    public void parse2List() {
        //TODO 校验       workbook的校验
        //TODO 解析表头（忽略掉不存在的）
        //TODO 解析行(先遍历，后赋值)
        //TODO 返回统一体
    }


    /**
     * workbook转换为java对象,默认忽略第一列
     *
     * @param workbook    workbook对象
     * @param targetClass 目标类型
     * @return List<T>转换后的对象
     */
    public static <T> List<T> book2listImpl(Workbook workbook, Class<T> targetClass, String... ignores) {

        Assert.notNull(workbook, "workbook must not be null!");
        Assert.notNull(targetClass, "targetClass must not be null!");
        Sheet sheetAt = workbook.getSheetAt(0);
        Assert.notNull(sheetAt, "模板格式不正确,位于第一页！");
        Row row = sheetAt.getRow(0);
        Assert.notNull(row, "模板格式不正确,位于第一行！");

        //表头解析
        Map<Integer, PoiCell> headerMapping = PoiCellParser.parseCellByRow(row);
        //字段解析
        Map<String, Field> allFieldMap = ParseBySwagger.getFiledMap(targetClass, ignores);


        //解析
        List<T> resultList = Lists.newArrayList();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int s = 0; s < numberOfSheets; s++) {
            Sheet sheetAtIndex = workbook.getSheetAt(s);
            List<T> parse = parse(targetClass, sheetAtIndex, allFieldMap, headerMapping);
            resultList.addAll(parse);
        }

        return resultList;
    }

    private static <T> List<T> parse(Class<T> targetClass, Sheet sheetAt,
                                     Map<String, Field> allFieldMap,
                                     Map<Integer, PoiCell> headerMapping) {
        List<T> resultList = Lists.newArrayList();
        int cellNum = 0, rowNum = 0;
        for (int i = 1; i <= sheetAt.getLastRowNum(); i++) {
            rowNum = i;
            T t = null;
            try {
                Row singleRow = sheetAt.getRow(i);
                t = targetClass.newInstance();
                for (int j = 0; j < singleRow.getLastCellNum(); j++) {
                    cellNum = j;
                    Cell cell = singleRow.getCell(j);

                    PoiCell cellValue;
                    try {
                        cellValue = PoiCellParser.getCellValue(cell);
                    } catch (Exception e) {
                        continue;
                    }

                    String s = headerMapping.get(j).getCellStrValue();
                    Field field = allFieldMap.get(s);

                    if (cellValue.getJavaType() != null && field != null) {
                        field.setAccessible(true);
                        Class<?> type = field.getType();
                        if (type.equals(cellValue.getJavaType())) {
                            field.set(t, cellValue.getValue());
                        } else {
                            Object value = cellValue.getValue();
                            if (value != null && StringUtils.isNotBlank(value.toString())) {
                                field.set(t, PoiCellParser.converter(value, type));
                            }
                        }
                    }
                }
                resultList.add(t);
            } catch (Exception e) {
//                e.printStackTrace();
                if (PoiConfig.IMPORT_VALID_WHOLE_EXCEL) {
                    throw new RuntimeException("模板格式不正确: 行 " + rowNum + ", 列 " + cellNum);
                } else {
                    resultList.add(t);
                    //TODO 记录错误行的数据，且抛出
                }
            }
        }
        return resultList;
    }
}
