package tw.com.seesawin.poiDemo.service;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class genExcelService {

	public void gen(XSSFWorkbook workbook) {
		System.out.println("generate excel...");

		XSSFSheet sheet = workbook.createSheet("sheet01");

		// 固定欄位寬度設定
		sheet.setColumnWidth(0, 256 * 20);
		sheet.setColumnWidth(1, 256 * 20);
		sheet.setColumnWidth(2, 256 * 20);

		XSSFRow title_header = sheet.createRow(0);

		XSSFCell title00 = title_header.createCell(0);
		title00.setCellValue("title01");

		XSSFCell title01 = title_header.createCell(1);
		title01.setCellValue("title02");

		XSSFCell title02 = title_header.createCell(2);
		title02.setCellValue("title03");

		XSSFCell cell = null;
		int initRow = 1;
		
		for (int i = 0; i < 10; i++) {
			XSSFRow detailRow = sheet.createRow(initRow + i);
			
			cell = detailRow.createCell(0);
			cell.setCellValue("111");
			
			cell = detailRow.createCell(1);
			cell.setCellValue("222");
			
			cell = detailRow.createCell(2);
			cell.setCellValue("333");
		}
	}

}
