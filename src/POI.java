import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class POI {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
           File file = new File("src//marketing.sheet.xlsx");
           FileInputStream fis = new FileInputStream(file);
           XSSFWorkbook workbook = new XSSFWorkbook(fis);
           XSSFSheet  sheet = workbook.getSheetAt(0);
          XSSFRow row = sheet.getRow(3);
          XSSFCell cell = row.getCell(0);
           String value = cell.getStringCellValue();
           System.out.println(value);
	}

}

