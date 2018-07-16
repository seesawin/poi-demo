package tw.com.seesawin.poiDemo.service;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class BaseAbstractService {

	protected Map<String, XSSFCellStyle> cellStyleMap;

	/**
	 * 設定style
	 * 
	 * @param workbook
	 */
	protected void setStyle(XSSFWorkbook workbook) {
		cellStyleMap = new HashMap<String, XSSFCellStyle>();

		XSSFCellStyle style_01 = workbook.createCellStyle();
		XSSFCellStyle style_02 = workbook.createCellStyle();

		cellStyleMap.put("style_01", style_01);
		cellStyleMap.put("style_02", style_02);

		// format 01
		XSSFFont font_01 = workbook.createFont();
		font_01.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font_01.setFontHeightInPoints((short) 14);
		style_01.setFont(font_01);
		style_01.setFont(font_01);
		style_01.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style_01.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style_01.setFillForegroundColor(HSSFColor.YELLOW.index);
		style_01.setFillPattern((short) 1);

		// format 02
		XSSFFont font_02 = workbook.createFont();
		font_02.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
		font_02.setFontHeightInPoints((short) 12);
		style_02.setFont(font_02);
		style_02.setAlignment(XSSFCellStyle.ALIGN_RIGHT);

	}

}
