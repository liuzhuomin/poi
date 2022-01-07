package ziv.excel.news.invoker;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * poi处理器，核心代码
 *
 * @param <T>
 * @author liuliuliu
 * @since 2021/1/20
 */
public class PoiIInvokerImpl<T> {
    /**
     * 一页最大
     */
    final int sheetMaxRow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
    /**
     * 大数据量
     */
    final int bigDataRow = sheetMaxRow;

    private Workbook book;

    /**
     * lit to book 实现
     *
     * @param sheetName sheetName
     * @param dataList  数据源
     * @param otherData 扩展数据（一般用于统计）
     * @param <S>
     * @return Workbook
     */
    public <S> Workbook generate(String sheetName,
                                 List<T> dataList,
                                 PoiInvoker<List<T>> dataListInvoker,
                                 List<S> otherData,
                                 PoiInvoker<List<S>> otherDataInvoker) {
        //获取数据
        Map<String, String> headers = dataListInvoker.getHeaders(dataList);
        List<Map<String, Object>> rows = dataListInvoker.getRows(dataList);
        PoiStyleInvoker poiStyleInvoker = dataListInvoker.getPoiStyleInvoker();

        //生成shell
        book = dataList.size() > bigDataRow ? new SXSSFWorkbook() : new HSSFWorkbook();
        Sheet sheet = book.createSheet(sheetName);

        CellStyle headerStyle = poiStyleInvoker.getHeaderStyle(book);
        CellStyle rowStyle = poiStyleInvoker.getRowStyle(book);
        dataing(headers, rows, sheet, headerStyle, rowStyle);

        //扩展行
        if (otherDataInvoker != null && !otherData.isEmpty()) {
            Map<String, String> otherHeaders = otherDataInvoker.getHeaders(otherData);
            List<Map<String, Object>> otherRows = otherDataInvoker.getRows(otherData);
            dataing(otherHeaders, otherRows, sheet, headerStyle, rowStyle);
        }
        //打开第一页sheet
        book.setActiveSheet(0);
        return book;
    }

    /**
     * 表格数据处理，如果某一页超过了最大限制，递归此函数
     *
     * @param headers     表头
     * @param rows        表行
     * @param sheet       sheet页
     * @param headerStyle 表头样式
     * @param rowStyle    表行样式
     */
    private void dataing(Map<String, String> headers, List<Map<String, Object>> rows, Sheet sheet, CellStyle headerStyle, CellStyle rowStyle) {
        List<String> indexList = new ArrayList<>(headers.size());
        int lastRowNum = sheet.getLastRowNum();
        //第一行表头
        Row headerRow = sheet.createRow(lastRowNum == 0 ? 0 : lastRowNum + 1);
        Set<String> strings = headers.keySet();
        for (String string : strings) {
            short lastCellNum = headerRow.getLastCellNum();
            int index = lastCellNum == -1 ? 0 : lastCellNum;
            Cell cell = headerRow.createCell(index);
            cell.setCellValue(headers.get(string));
            cell.setCellStyle(headerStyle);
            indexList.add(string);
        }
        //常规行
        List<Map<String, Object>> row = createRow(rows, indexList, rowStyle, sheet);
        if (row != null && !row.isEmpty()) {
            dataing(headers, row, book.getSheetAt(book.getActiveSheetIndex()), headerStyle, rowStyle);
        }
    }

    /**
     * 根据rows对象动态创建行
     *
     * @param rows      行集合
     * @param indexList
     * @param rowStyle  行样式
     * @param sheet     sheet页码
     */
    private List<Map<String, Object>> createRow(List<Map<String, Object>> rows, List<String> indexList, CellStyle rowStyle, Sheet sheet) {
        int size = rows.size();
        for (int i = 0; i < size; i++) {

            Map<String, Object> objectMap = rows.get(i);

            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum >= sheetMaxRow) {
                String format = String.format("%s(%s)", book.getSheetAt(0).getSheetName(), book.getActiveSheetIndex() + 1);
                book.createSheet(format);
                book.setActiveSheet(book.getActiveSheetIndex() + 1);
                return rows.subList(i, size);
            }

            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            for (String str : indexList) {
                Object rowValue = objectMap.get(str);
                short lastCellNum = row.getLastCellNum();
                Cell cell = row.createCell(lastCellNum == -1 ? 0 : lastCellNum);
                cell.setCellValue(rowValue == null ? null : rowValue.toString());
                cell.setCellStyle(rowStyle);
            }
        }
        return null;
    }


}
