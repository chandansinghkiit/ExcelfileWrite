package com.mystyle;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

@Component
public class TestExcelfilewrite {
	
	@Autowired
	JdbcTemplate jdbcTemplate;
	
	public String ExcuteWithParam()
	
	{
		
		List<Map<String, Object>> ListofRowLines = jdbcTemplate.queryForList(" select "
				 +"id ,"+
				 "first_name ,"+
				 "surname ,"+
				 "Dob ,"+
				 "Email ,"+
				 "Telephone ,"+
				 "Address ,"+
				 "Image ,"+
				 "Gender ,"+
				 "Address2 ,"+
				 "Apartment ,"+
				 "Post_code ,"+
				 "course_id "+
			
				 " from student_information");
		
		System.out.println(ListofRowLines);
		
		
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("StudentInfromation Data");
		/*
		 * SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
		 * // Condition 1: Formula Is =A2=A1 (White Font) ConditionalFormattingRule
		 * rule1 = sheetCF.createConditionalFormattingRule("ROW(),0)");
		 * PatternFormatting fill1 = rule1.createPatternFormatting();
		 */
        CellStyle backgroundStyle = workbook.createCellStyle();

        backgroundStyle.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        backgroundStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
      
		
		if (ListofRowLines != null && !ListofRowLines.isEmpty()) {
			int rownum=0;
			for (Map<String, Object> row : ListofRowLines) {
			    Row row1 = sheet.createRow(rownum++);
			    int cellnum = 0;
				for (Iterator<Map.Entry<String, Object>> it = row.entrySet().iterator(); it.hasNext();) {
					Map.Entry<String, Object> entry = it.next();
					String key = entry.getKey();
					Object obj = entry.getValue();
					
				      Cell cell = row1.createCell(cellnum++);
		               if(rownum==1)
		               {
		            	   cell.setCellStyle(backgroundStyle);
		            	    if(key instanceof String) {
			                    cell.setCellValue((String)key);
		            	    }else {
		            	        cell.setCellValue(key);
		            	    }
		            	    
		            	    
		
		               }else {
		            	    if(obj instanceof String)
			                    cell.setCellValue((String)obj);
			                else if(obj instanceof Integer)
			                    cell.setCellValue((Integer)obj); 
		               }
		           
		            
				}
			}
			
			 try
		        {
		        	  
		            //Write the workbook in file system
		            FileOutputStream out = new FileOutputStream(new File("excelfile_chandan.xlsx"));
		          
		            workbook.write(out);
		            out.close();
		            System.out.println("excelfile written successfully on disk.");
		        }
		        catch (Exception e)
		        {
		            e.printStackTrace();
		        }	
			
			
			
			
		}
		
		return "student";
		
	}
	
	
	
}
