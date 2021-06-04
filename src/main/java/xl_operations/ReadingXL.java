package xl_operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingXL {
	public static void main(String[] args) {
		String excelFilePath = "./country.xlsx";
		FileInputStream inputStream;
		try {
			inputStream = new FileInputStream(excelFilePath);
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			
			//using for loop
			
			int rows=sheet.getLastRowNum();
			int cols=sheet.getRow(1).getLastCellNum();
			
			for(int r=0;r<=rows;r++)
			{XSSFRow row=  sheet.getRow(r);
				for(int c=0;c<cols;c++)
				{
				XSSFCell cell=	row.getCell(c);
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()+ " ");break;
				case NUMERIC: System.out.println(cell.getNumericCellValue()+ " ");break;
				case BOOLEAN: System.out.println(cell.getBooleanCellValue()+ " ");break;
				
				}
				}
				System.out.println();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
