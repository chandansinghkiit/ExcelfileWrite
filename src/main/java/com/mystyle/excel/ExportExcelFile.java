package com.mystyle.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class ExportExcelFile extends ExcelExportUtility<Object> {

	public String exportExcelFile(String[] strings, List<Map<String, Object>> listofRowLines) {

	
		String excelFileName = "testexcelfile.xlsx";
		// Write the workbook in file system

		try {
			FileOutputStream out = new FileOutputStream(new File(excelFileName));

			SXSSFWorkbook wb = exportExcel(strings, listofRowLines);

			wb.write(out);
			out.flush();
			wb.dispose();
			wb.close();
			out.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return excelFileName;
	}

}