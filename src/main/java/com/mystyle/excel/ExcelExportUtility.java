package com.mystyle.excel;

import java.math.BigDecimal;
import java.sql.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public abstract class ExcelExportUtility<E extends Object> {

	protected SXSSFWorkbook wb;
	protected Sheet sh;
	protected static final String EMPTY_VALUE = " ";

	/**
	 * This method demonstrates how to Auto resize Excel column
	 */
	private void autoResizeColumns(int listSize) {

		for (int colIndex = 0; colIndex < listSize; colIndex++) {
			sh.autoSizeColumn(colIndex);
		}
	}

	/**
	 * 
	 * This method will return Style of Header Cell
	 * 
	 * @return
	 */

	protected CellStyle getHeaderStyle() {
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		Font font = wb.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		style.setFont(font);
		

		return style;
	}

	/**
	 * 
	 * This method will return style for Normal Cell
	 * 
	 * @return
	 */

	protected CellStyle getNormalStyle() {
		CellStyle style = wb.createCellStyle();

		return style;
	}

	/**
	 * @param columns
	 */
	private void fillHeader(String[] columns) {
		wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
		sh = wb.createSheet("Validated Data");
		CellStyle headerStle = getHeaderStyle();

		for (int rownum = 0; rownum < 1; rownum++) {
			Row row = sh.createRow(rownum);

			for (int cellnum = 0; cellnum < columns.length; cellnum++) {
				Cell cell = row.createCell(cellnum);
				cell.setCellValue(columns[cellnum]);
				cell.setCellStyle(headerStle);
			}

		}
	}

	/**
	 * @param columns
	 * @param dataList
	 * @return
	 */
	public final SXSSFWorkbook exportExcel(String[] columns, List<Map<String, Object>> ListofRowLines) {

		fillHeader(columns);
		fillData(ListofRowLines);
		//autoResizeColumns(columns.length);

		return wb;
	}

	/**
	 * @param dataList
	 * @return 
	 */
	 void fillData(List<Map<String, Object>> ListofRowLines) {
		 
			if (ListofRowLines != null && !ListofRowLines.isEmpty()) {
				int rownum=1;
				for (Map<String, Object> row : ListofRowLines) {
				    Row row1 = sh.createRow(rownum++);
				    int cellnum = 0;
					for (Iterator<Map.Entry<String, Object>> it = row.entrySet().iterator(); it.hasNext();) {
						Map.Entry<String, Object> entry = it.next();
						String key = entry.getKey();
						Object obj = entry.getValue();
						
					      Cell cell = row1.createCell(cellnum++);
					    		  
			            	    if(obj instanceof String)
				                    cell.setCellValue((String)obj);
				                else if(obj instanceof Integer)
				                    cell.setCellValue((Integer)obj);
				                else if(obj instanceof BigDecimal)
				                	 cell.setCellValue((RichTextString)obj);
				                else if(obj instanceof Double)
				                	 cell.setCellValue((Double)obj);
				                else if(obj instanceof Date)
				                	 cell.setCellValue((Date)obj);
				                else if(obj instanceof Number)
				                	 cell.setCellValue((double)obj);
			            
			           
			            
					}
				}
	 }


}
}
