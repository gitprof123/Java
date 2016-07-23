package excelAndXML;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//List <String> output = new ArrayList<String>();
				
		String filePath = "./src/Inputs/UserNameAndPasswords.xlsx";
		String sheetName="UserName";
		
		Workbook excelWorkbook = null;
		String fileExtensionName = filePath.substring(filePath.lastIndexOf("."));
		FileInputStream inputStream = new FileInputStream(filePath);
		
		if(fileExtensionName.equalsIgnoreCase(".xls"))
		{
			excelWorkbook = new HSSFWorkbook(inputStream);
		}
		else if(fileExtensionName.equalsIgnoreCase(".xlsx"))
		{
			excelWorkbook = new XSSFWorkbook(inputStream);
		}
		
		Sheet excelWorkSheet = excelWorkbook.getSheet(sheetName);
		int rowCount = excelWorkSheet.getLastRowNum();
		Object[][] output = new Object[rowCount][];
		for(int i = 0; i<rowCount; i++)
		{
			Row excelRow = excelWorkSheet.getRow(i);
			List<Object> rowList = new ArrayList<Object>();
			int colCount= excelRow.getLastCellNum();
			for(int j=0;j<colCount;j++)
			{
				rowList.add(excelRow.getCell(j));
			}
			
			output[i]= rowList.toArray();
		}
		
		for(int i =0; i<output.length;i++)
		{
			Object[] user = output[i];
			for(int j=0;j<user.length;j++)
			{
				System.out.print(user[j]+" ");
			}
			System.out.println();
		}
	}

}
