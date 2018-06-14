package tw.com.seesawin.poiDemo.main;

import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import tw.com.seesawin.poiDemo.service.genExcelService;
import tw.com.seesawin.util.CwFileUtils;

public class Entry {

	public static void main(String[] args) throws Exception {
		System.out.println("start...");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		genExcelService service = new genExcelService();
		service.gen(workbook);
		
		String reportName = "seesawin_" + new Date().getTime();
		String destination = "C:\\test";
		
		CwFileUtils.createExcelFile(workbook, reportName, destination);
	}

}
