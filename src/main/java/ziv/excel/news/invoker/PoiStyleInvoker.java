package ziv.excel.news.invoker;


import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 样式执行器
 *
 * @author liuliuliu
 * @since 2021/10/28
 */
public interface PoiStyleInvoker {

    /**
     * 第一行标题单元样式（生成excel时会在第一行的每一列使用此样式）
     * @param workbook workbook对象
     * @return 每一列的样式
     */
    CellStyle getHeaderStyle(Workbook workbook);

    /**
     * 数据行的单元样式（生成excel时会在除了第一行的每一列使用此样式）
     * @param workbook workbook对象
     * @return 每一列的样式
     */
    CellStyle getRowStyle(Workbook workbook);
}
