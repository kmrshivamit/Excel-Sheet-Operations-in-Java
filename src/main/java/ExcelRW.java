import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.SSCellRange;

public class ExcelRW {
	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static Row row;
	private static Cell cell;
	public static void main(String[] args) {
		try {
			fis=new FileInputStream("./testdata.xlsx");
			System.out.println(fis.getFD());;
		
			wb=WorkbookFactory.create(fis);
			sh=wb.getSheet("Sheet1");
			int noOfRows=sh.getLastRowNum();
			System.out.println(noOfRows);
			
			row=sh.createRow(0);
			cell=row.createCell(0);
			cell.setCellValue("QAV");
			System.out.println(cell.getStringCellValue());
			fos=new FileOutputStream("./testdata.xlsx");
			wb.write(fos);;
			fos.flush();
			fos.close();
			System.out.println("done");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	

}
