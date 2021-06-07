package xl_operations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingXL {
	public static void main(String[] args) {
		System.out.println("program to write things on xl sheet");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		Object empdata[][] = { { "EmpId", "Name", "Job" }, { 101, "David", "Engineer" }, { 102, "Smith", "Manager" },
				{ 103, "David", "Engineer" } };
		
		//using for loop
		//Just for example
		
		int rows=empdata.length;
		int cols=empdata[0].length;
		
		System.out.println(rows);
		System.out.println(cols);
		
		for(int r=0;r<rows;r++) {
			XSSFRow row=sheet.createRow(r);
			for(int c=0;c<cols;c++) {
				XSSFCell cell=row.createCell(c);
				Object value=empdata[r][c];
				if(value instanceof String)
					cell.setCellValue((String)value);
				 if(value instanceof Integer)
					 cell.setCellValue((Integer) value);
				 if(value instanceof Boolean)
					 cell.setCellValue((Boolean) value);
				 
			}
		}
		String filePath="./employee.xlsx";
		try {
			FileOutputStream outstream=new FileOutputStream(filePath);
			workbook.write(outstream);
			outstream.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
	}

}
