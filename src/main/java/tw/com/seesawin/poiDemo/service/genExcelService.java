package tw.com.seesawin.poiDemo.service;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class genExcelService extends BaseAbstractService {

	public void gen(XSSFWorkbook workbook) {
		System.out.println("generate excel...");

		// 設定樣式
		this.setStyle(workbook);

		XSSFSheet sheet = workbook.createSheet("sheet01");

		// 凍結標題欄位
		sheet.createFreezePane(1, 2);

		// 篩選
		sheet.setAutoFilter(CellRangeAddress.valueOf("A2:G2"));

		// 固定欄位寬度設定
		sheet.setColumnWidth(0, 256 * 20);
		sheet.setColumnWidth(1, 256 * 20);
		sheet.setColumnWidth(2, 256 * 20);
		sheet.setColumnWidth(3, 256 * 20);
		sheet.setColumnWidth(4, 256 * 20);
		sheet.setColumnWidth(5, 256 * 20);
		sheet.setColumnWidth(6, 256 * 20);

		// 合併儲存格設置
		CellRangeAddress cellRangeRoop = null;
		cellRangeRoop = new CellRangeAddress(0, 1, 0, 0);
		sheet.addMergedRegion(cellRangeRoop);
		cellRangeRoop = new CellRangeAddress(0, 1, 1, 1);
		sheet.addMergedRegion(cellRangeRoop);
		cellRangeRoop = new CellRangeAddress(0, 1, 2, 2);
		sheet.addMergedRegion(cellRangeRoop);
		cellRangeRoop = new CellRangeAddress(0, 0, 3, 4);
		sheet.addMergedRegion(cellRangeRoop);
		cellRangeRoop = new CellRangeAddress(0, 0, 5, 6);
		sheet.addMergedRegion(cellRangeRoop);

		// 第一排標題
		XSSFRow title_header = sheet.createRow(0);

		XSSFCell title00 = title_header.createCell(0);
		title00.setCellValue("title01");
		title00.setCellStyle(cellStyleMap.get("style_01"));

		XSSFCell title01 = title_header.createCell(1);
		title01.setCellValue("title02");
		title01.setCellStyle(cellStyleMap.get("style_01"));

		XSSFCell title02 = title_header.createCell(2);
		title02.setCellValue("title03");
		title02.setCellStyle(cellStyleMap.get("style_01"));

		XSSFCell title03 = title_header.createCell(3);
		title03.setCellValue("title04");
		title03.setCellStyle(cellStyleMap.get("style_01"));

		XSSFCell title04 = title_header.createCell(5);
		title04.setCellValue("title05");
		title04.setCellStyle(cellStyleMap.get("style_01"));
		
		// 第二排標題
		XSSFRow title_header2 = sheet.createRow(1);

		XSSFCell title03_1 = title_header2.createCell(3);
		title03_1.setCellValue("title03_1");
		title03_1.setCellStyle(cellStyleMap.get("style_01"));
		XSSFCell title03_2 = title_header2.createCell(4);
		title03_2.setCellValue("title03_2");
		title03_2.setCellStyle(cellStyleMap.get("style_01"));
		XSSFCell title04_1 = title_header2.createCell(5);
		title04_1.setCellValue("title04_1");
		title04_1.setCellStyle(cellStyleMap.get("style_01"));
		XSSFCell title04_2 = title_header2.createCell(6);
		title04_2.setCellValue("title04_2");
		title04_2.setCellStyle(cellStyleMap.get("style_01"));

		XSSFCell cell = null;
		int initRow = 2;

		for (int i = 0; i < 50; i++) {
			XSSFRow detailRow = sheet.createRow(initRow + i);

			cell = detailRow.createCell(0);
			cell.setCellValue("111" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(1);
			cell.setCellValue("222" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(2);
			cell.setCellValue("33344445555" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(3);
			cell.setCellValue("aaaaa" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(4);
			cell.setCellValue("bbbbb" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(5);
			cell.setCellValue("ccccc" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));

			cell = detailRow.createCell(6);
			cell.setCellValue("ddddd" + i);
			cell.setCellStyle(cellStyleMap.get("style_02"));
		}
	}

}
