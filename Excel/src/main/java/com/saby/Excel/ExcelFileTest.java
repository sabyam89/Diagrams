package com.saby.Excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileTest {
 
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Excel");
        ArrayList<Object> bookData = new ArrayList<Object>();
        try
		{
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection con=DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe", "system", "saby1");
			PreparedStatement prepstmt = con.prepareStatement("Select * from RAW_REPORT");
			ResultSet rs = null;
			rs = prepstmt.executeQuery();
			ArrayList<Object> abookData = new ArrayList<Object>();
			abookData.add("ID");
			abookData.add("SALES");
			abookData.add("COMMISON");
			abookData.add("STAFFNAME");
			abookData.add("DATE_SALES");
			bookData.add(abookData);
			while(rs.next()){
				ArrayList<Object> abookData1 = new ArrayList<Object>();
				abookData1.add(rs.getInt("ID"));
				abookData1.add(rs.getString("SALES"));
				abookData1.add(rs.getInt("COMMISON"));
				abookData1.add(rs.getString("STAFFNAME"));
				abookData1.add(rs.getString("DATE_SALES"));
				bookData.add(abookData1);
			}
		}
		catch(Exception connection)
		{
			connection.printStackTrace();
		}
 
        int rowCount = 0;
        
        for (Object aookData : bookData) {
            Row row = sheet.createRow(++rowCount); 
            int columnCount = 0;
            int i = ((ArrayList<Object>) aookData).size();
            for (int j =0;j<i;j++) {
                Cell cell = row.createCell(++columnCount);
                if(rowCount==1){
                	CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
                    Font font = sheet.getWorkbook().createFont();
                    font.setBold(true);
                    font.setFontHeightInPoints((short) 16);
                    cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                    cellStyle.setFont(font);
                    cell.setCellStyle(cellStyle);
                }
                if (((ArrayList<Object>) aookData).get(j) instanceof String) {
                    cell.setCellValue((String) ((ArrayList<Object>) aookData).get(j));
                } else if (((ArrayList<Object>) aookData).get(j) instanceof Integer) {
                    cell.setCellValue((Integer) ((ArrayList<Object>) aookData).get(j));
                }
            }
             
        }
        try (FileOutputStream outputStream = new FileOutputStream("ExcelTest.xlsx")) {
            workbook.write(outputStream);
            System.out.println("Done");
        }
    }
 
}