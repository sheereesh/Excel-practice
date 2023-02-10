package highestNOInJava;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class highestNumber {

	public static void main(String[] args) throws IOException {
		//FileInputStream fis=new FileInputStream("C:\\Users\\sireesh\\Desktop\\demo\\student.xlsx");
		
		//FileInputStream fis =new FileInputStream("C://Users//sireesh//Desktop//demo//student.xlsx");
		FileInputStream fis =new FileInputStream(System.getProperty("user.dir")+"/src/test/resources/properties/rajesh.xlsx");
		//FileInputStream fis =new FileInputStream(System.getProperty("user.dir")+"/src/test/resources/student.xlsx");
		//System.out.println(fis);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		 XSSFSheet sheet=workbook.getSheet("student data");
	int	rows =sheet.getLastRowNum();
	int cols=sheet.getRow(0).getLastCellNum();
	System.out.println(rows+"number of rows "+cols+"number of cols");
String name=sheet.getRow(2).getCell(1).getStringCellValue();
		System.out.println(name);
		double value=sheet.getRow(3).getCell(3).getNumericCellValue();
		System.out.println(value);
	}

}
