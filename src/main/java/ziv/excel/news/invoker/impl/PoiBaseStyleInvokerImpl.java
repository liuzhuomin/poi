package ziv.excel.news.invoker.impl;

import ziv.excel.news.invoker.PoiStyleInvoker;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

/**
 * 基础样式实现
 * @author liuliuliu
 * @since 2021/10/29
 */
public class PoiBaseStyleInvokerImpl implements PoiStyleInvoker {

	@Override
	public CellStyle getHeaderStyle(Workbook book) {
		CellStyle cellStyle = book.createCellStyle();
//		cellStyle.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
		cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getIndex());
		setUsual(cellStyle, BorderStyle.MEDIUM);
		//字体
		Font endHeaderStyleFont = book.createFont();
		endHeaderStyleFont.setBold(true);
		cellStyle.setFont(endHeaderStyleFont);
		return cellStyle;
	}

	@Override
	public CellStyle getRowStyle(Workbook book) {
		CellStyle cellStyle = book.createCellStyle();
//		cellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
		cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex2());
		setUsual(cellStyle,BorderStyle.THIN);
		return cellStyle;
	}
	/**
	 * 设置cell的边框以及边框样式，换行，居中
	 *
	 * @param endHeaderStyle cell样式对象
	 * @param thin           边框格式对象
	 */
	private static void setUsual(CellStyle endHeaderStyle, BorderStyle thin) {

		endHeaderStyle.setBorderBottom(thin);
		endHeaderStyle.setBorderLeft(thin);
		endHeaderStyle.setBorderRight(thin);
		endHeaderStyle.setBorderTop(thin);

//        endHeaderStyle.setBottomBorderColor(HSSFColor.BLUE.index);
//        endHeaderStyle.setLeftBorderColor(HSSFColor.BLUE.index);
//        endHeaderStyle.setRightBorderColor(HSSFColor.BLUE.index);
//        endHeaderStyle.setTopBorderColor(HSSFColor.BLUE.index);

		endHeaderStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		endHeaderStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		endHeaderStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
		endHeaderStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());


		//居中
		endHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
		endHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		//自动换行
		endHeaderStyle.setWrapText(true);
		//背景填充
		endHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

}
