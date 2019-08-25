package com.qnb.message.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.qnb.message.entity.User;

@Component
public class Util {

	private static final String FILE_NAME = "qnb.xlsx";
	private static String[] columns = { "Kayıt No", "Ad Soyad", "Mesaj", "Sicil No", "Kayıt Tarihi" };

	public String createXLSFile(User user) throws FileNotFoundException, IOException {
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		Date date = new Date();
		String cwd = "/";
		File f = new File(cwd + FILE_NAME);
		if (f.exists() && !f.isDirectory()) {
			XSSFWorkbook workbook = null;
			XSSFSheet sheet;
			FileInputStream file = new FileInputStream(new File(cwd + "/" + FILE_NAME));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
			int rowCount = sheet.getLastRowNum() + 1;
			Row empRow = sheet.createRow(rowCount);
			Cell c0 = empRow.createCell(0);
			c0.setCellValue(rowCount);
			Cell c1 = empRow.createCell(1);
			c1.setCellValue(user.getName());
			Cell c2 = empRow.createCell(2);
			c2.setCellValue(user.getMessage());
			Cell c3 = empRow.createCell(3);
			c3.setCellValue(user.getSicil());
			Cell c4 = empRow.createCell(4);
			c4.setCellValue(formatter.format(date));
			FileOutputStream out = new FileOutputStream(new File(cwd + FILE_NAME));
			workbook.write(out);
			out.close();
			System.out.println("Kayit Eklendi");
			workbook.close();
			return "redirect:/success";
		} else {
			Workbook workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet("Kayit");

			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 14);
			headerFont.setColor(IndexedColors.RED.getIndex());

			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);
			Row headerRow = sheet.createRow(0);

			for (int i = 0; i < columns.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headerCellStyle);
			}
			Row empRow = sheet.createRow(1);
			Cell c0 = empRow.createCell(0);
			c0.setCellValue(1);
			Cell c1 = empRow.createCell(1);
			c1.setCellValue(user.getName());
			Cell c2 = empRow.createCell(2);
			c2.setCellValue(user.getMessage());
			Cell c3 = empRow.createCell(3);
			c3.setCellValue(user.getSicil());
			Cell c4 = empRow.createCell(4);
			c4.setCellValue(formatter.format(date));

			for (int i = 0; i < columns.length; i++) {
				sheet.autoSizeColumn(i);
			}

			FileOutputStream fileOut = new FileOutputStream(cwd + FILE_NAME);
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();

			return "redirect:/success";
		}
	}

	public Object getAllMessages() {
		
		List<User> userList = new ArrayList<>();
		String cwd = "/";
		// Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(new File(cwd + FILE_NAME));
		} catch (EncryptedDocumentException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (InvalidFormatException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");


        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

           
           
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                Cell cell2 = cellIterator.next();
                String cellValue2 = dataFormatter.formatCellValue(cell2);
                Cell cell3 = cellIterator.next();
                String cellValue3 = dataFormatter.formatCellValue(cell3);
                Cell cell4 = cellIterator.next();
                String cellValue4 = dataFormatter.formatCellValue(cell4);
                Cell cell5 = cellIterator.next();
                String cellValue5 = dataFormatter.formatCellValue(cell5);
                User u = new User();
                u.setId(cellValue);
                u.setName(cellValue2);
                u.setMessage(cellValue3);
                u.setSicil(cellValue4);
                u.setDate(cellValue5);
                userList.add(u); 
        }


        // Closing the workbook
        try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
		return userList;
    }
	}
