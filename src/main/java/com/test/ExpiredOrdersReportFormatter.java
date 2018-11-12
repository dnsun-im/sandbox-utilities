package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAdjusters;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row; 

public class ExpiredOrdersReportFormatter {
	public static final int DOB_INDEX = 2;
	public static final int EXPIRE_DATE_INDEX = 5;
	
	private static final DateTimeFormatter EURO_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");
	private static final DateTimeFormatter EXPIRED_FORMAT = DateTimeFormatter.ofPattern("MMM d yyyy h:mma");
	private static final DateTimeFormatter TARGET_FORMAT = DateTimeFormatter.ofPattern("MMM dd, yyyy");
	
	private LocalDate date;
	private String reportPath;
	private String targetDirectory;

	LocalDate end = LocalDate.now().plusDays(45);
	
	public static void main(String[] args) throws FileNotFoundException, IOException { 
		LocalDate date = LocalDate.parse(args[0], EURO_FORMAT);
		String reportPath = args[1];
		String targetDirectory = args[2];
		
		ExpiredOrdersReportFormatter formatter = new ExpiredOrdersReportFormatter(date, reportPath, targetDirectory);
		formatter.format();
	}
	
	public ExpiredOrdersReportFormatter(LocalDate date, String reportPath, String targetDirectory) {
		this.date = date;
		this.reportPath = reportPath;
		this.targetDirectory = targetDirectory;
	}
	
	public void format() throws FileNotFoundException, IOException {
		HSSFWorkbook workbook = getWorkbook();
		HSSFSheet sheet = workbook.getSheetAt(0);
		int rcount = sheet.getPhysicalNumberOfRows();
		
		for(int i=2; i < rcount; i++) {
			Row row = sheet.getRow(i);
			if(shouldDelete(row)) {
				sheet.removeRow(row);
				sheet.shiftRows(i+1, rcount, -1);
				i--;
			}
			else {
				String dob = getFormattedDob(row);
				row.getCell(DOB_INDEX, Row.CREATE_NULL_AS_BLANK).setCellValue(dob);
			}
		} 
		
		FileOutputStream outputStream = new FileOutputStream(targetDirectory + "/target.xls");
		workbook.write(outputStream);
	}
	
	private boolean shouldDelete(Row row) {
		Cell expiration = row.getCell(EXPIRE_DATE_INDEX, Row.CREATE_NULL_AS_BLANK);
		boolean should = true;
		if(expiration == null || expiration.getStringCellValue().equals("")) { 
			should = false;
		}
		else {
			String expText = expiration.getStringCellValue().replace("  ", " ");
			LocalDate expiredDate = LocalDate.parse(expText, EXPIRED_FORMAT);
			
			LocalDate start = date.with(TemporalAdjusters.firstDayOfNextMonth());
			LocalDate end = start.with(TemporalAdjusters.lastDayOfMonth());
			should =  expiredDate.isAfter(end) || expiredDate.isBefore(start);
		} 
		return should;
	}
	
	private HSSFWorkbook getWorkbook() throws FileNotFoundException, IOException {
		HSSFWorkbook workbook = null;
		try(FileInputStream excelFile = new FileInputStream(new File(reportPath))) { 
			workbook = new HSSFWorkbook(excelFile); 
		}
		return workbook;
	}
	
	private String getFormattedDob(Row row) {
		String formatted = "";
		String dob = row.getCell(DOB_INDEX, Row.CREATE_NULL_AS_BLANK).getStringCellValue().trim();
		
		if("".equals(dob)) {
			return formatted;
		}
		
		try {
			formatted = LocalDate.parse(dob, EURO_FORMAT).format(TARGET_FORMAT); 
		}
		catch(DateTimeParseException e) {
			// ignore return empty space
			System.out.println("error parsing to date: " + dob);
		} 
		return formatted;
	} 
}
