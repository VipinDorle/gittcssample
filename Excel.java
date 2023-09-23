package firstdemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void SendResult(String user, String env, boolean result) {
		// TODO Auto-generated method stub
		FileInputStream file = null;
		int row=15;
		int col=17;
		int ResultRow=0;
		int ResultCol=0;
		XSSFRow XLrow=null;
		XSSFCell XLcell=null;
		
		
	
		XSSFWorkbook workbook = null;
		try {
			file = new FileInputStream("src/main/resources/Users.xlsx");
			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet= workbook.getSheetAt(0);
			user=Reformat(user);
			env=Reformat(env);
			
			
			for(int i=0;i<row;i++) {
				XLrow=sheet.getRow(i);
				for(int j=0;j<col;j++) {
					try{
					XLcell=XLrow.getCell(j);
					String CellValue=XLcell.toString();
					CellValue=Reformat(CellValue);
					System.out.println(CellValue+" "+i+" "+j);
					if(CellValue.equals(user)) {
						ResultRow=i;
						System.out.println("Found User");
					}
					if(CellValue.equals(env)) {
						ResultCol=j;
						System.out.println("Found Env");
					}
					file.close();
					}catch(NullPointerException p) {
						
					}
						
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {	
			
		}
		
		
		System.out.println("Result location"+ResultRow+" "+ResultCol);
		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream("src/main/resources/Users.xlsx");	
			XSSFSheet sheet= workbook.getSheetAt(0);
			XLrow=sheet.getRow(ResultRow);
			XLcell=XLrow.getCell(ResultCol);
			XLcell=XLrow.createCell(ResultCol,CellType.STRING);
			
			if(result)
			XLcell.setCellValue("Pass");
			else
			XLcell.setCellValue("Fail");
			System.out.println(XLcell.getStringCellValue());
			workbook.write(outputStream);
			workbook.close();
			
			outputStream.close();
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	public static String Reformat(String text) {
		text=text.toLowerCase();
		text=text.replace(" ", "");
		return text;
	}
	public static void FillSheet() {
		FileInputStream file = null;
		int row=17;
		int col=15;
		int ResultRow=0;
		int ResultCol=0;
		XSSFRow XLrow=null;
		XSSFCell XLcell=null;
		
		
	
		XSSFWorkbook workbook = null;
		try {
			
			file = new FileInputStream("src/main/resources/Users.xlsx");
			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet= workbook.getSheetAt(0);
			
			
			for(int i=1;i<row;i++) {
				XLrow=sheet.getRow(i);
				for(int j=3;j<col;j++) {
					try{
					XLcell=XLrow.getCell(j);
					XLcell=XLrow.createCell(j);
					XLcell.setCellValue("b");
					
					file.close();
					}catch(NullPointerException p) {
						
					}
						
				}
			}
			FileOutputStream outputStream = new FileOutputStream("src/main/resources/Users.xlsx");	
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {	
			
		}
		
	}

}
