package ApachePOI;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dtfeeding_rescomp {
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static FileInputStream flp;
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
//		String value =getdata(0,0);
//		System.out.println(value);
//		String value1 = getdata(1,0);
//		System.out.println(value1);
        for (int i=0; i<=4; i++){
        	for (int j=0; j<=1; j++){
        		if (j==0){
        			String value = getdata(i,j);
        			System.out.println(value);
        		}
        		else if(j==1){
        			int rollnum = getintdata(i,j);
        			System.out.println(rollnum);
        		}
        		
        		
        	}
        }
	}
	
	/**
	 * @param rownum
	 * @param col
	 * @return
	 * @throws IOException
	 */
	public static String getdata(int rownum, int col) throws IOException
	{
		flp= new FileInputStream("C:\\Data.xlsx");
		wb = new XSSFWorkbook(flp);
		sheet = wb.getSheet("LoginData");
		row = sheet.getRow(rownum);
		cell = row.getCell(col);
		return cell.getStringCellValue();
	}

	public static int getintdata(int rownum, int col) throws IOException
	{
		flp= new FileInputStream("C:\\Data.xlsx");
		wb = new XSSFWorkbook(flp);
		sheet = wb.getSheet("LoginData");
		row = sheet.getRow(rownum);
		cell = row.getCell(col);
		return (int) cell.getNumericCellValue();
	}
}
