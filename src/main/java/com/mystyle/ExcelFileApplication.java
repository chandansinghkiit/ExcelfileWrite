package com.mystyle;

import java.text.SimpleDateFormat;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;


@SpringBootApplication
public class ExcelFileApplication {

	@Autowired
	static TestExcelfilewrite externalFlatFileWatchService;

	public static void main(String[] args) {


		ConfigurableApplicationContext context = SpringApplication.run(ExcelFileApplication.class, args);
		externalFlatFileWatchService = context.getBean(TestExcelfilewrite.class);
		String str=externalFlatFileWatchService.ExcuteExcel();
		System.out.println("write program");
	}
	
	


		

}
