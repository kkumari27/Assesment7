package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Two {//declaring variables
	public static String vsearch;
	public static int xlRows;
	public static int xlCols;
	public static String xData[][];
	
	public static void main(String[] args) throws Exception {
//		TODO  Auto-generated method stub
		xlRead("C:\\Users\\kkumari27\\Desktop\\YahooDDF.xls");
		
		for(int i=0;i<xlRows;i++) {
//			wherever Execute=Y ,list it
			if(xData[i][1].equals("Y")) {	
		vsearch =xData[i][0];
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\kkumari27\\Downloads\\seleniumjar\\chromedriver.exe");
		WebDriver myD=new ChromeDriver();
		myD.get("https://www.yahoo.com/");
		Thread.sleep(2000);
		myD.findElement(By.xpath("//*[@id=\"uh-search-box\"]")).sendKeys(vsearch);
		Thread.sleep(2000);
		myD.findElement(By.xpath("//*[@id=\"uh-search-button\"]")).click();
		Thread.sleep(2000);
		System.out.println(myD.getTitle());
		 xData[i][2]=myD.getTitle();
		 System.out.println(myD.getTitle());
		myD.close();
		}
//			to write elements in the xlsheet
		xlwrite("C:\\Users\\kkumari27\\Desktop\\YahooDDF1.xls",xData);	
}
	}
	public static void xlRead(String sPath) throws Exception
	{
		File myFile=new File(sPath);
		FileInputStream myStream=new FileInputStream(myFile);
		HSSFWorkbook myworkbook=new HSSFWorkbook(myStream);
		HSSFSheet mySheet=myworkbook.getSheetAt(0);
		xlRows=mySheet.getLastRowNum()+1;
		xlCols=mySheet.getRow(0).getLastCellNum();
		xData=new String[xlRows][xlCols];
		for(int i=0;i<xlRows;i++)
		{
			HSSFRow row=mySheet.getRow(i);
			for(short j=0;j<xlCols;j++)
			{
				HSSFCell cell=row.getCell(j);
				String value=cellToString(cell);
				xData[i][j]=value;
				System.out.print("-"+xData[i][j]);
			}
			System.out.println();
		}
	}
		public static String cellToString(HSSFCell cell)
		{
			int type=cell.getCellType();
			Object result;
			switch(type)
			{
			case HSSFCell.CELL_TYPE_NUMERIC:
			result=cell.getNumericCellValue();
			break;
			case HSSFCell.CELL_TYPE_STRING:
			result=cell.getStringCellValue();
			break;
			case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
			case HSSFCell.CELL_TYPE_BLANK:
			result="-";
			case HSSFCell.CELL_TYPE_BOOLEAN:
			result=cell.getBooleanCellValue();
			case HSSFCell.CELL_TYPE_ERROR:
			result="This cell has some error";
			default:
			throw new RuntimeException("We do not support this cell type");
			}
			return result.toString();
			
		}
		
		public static void xlwrite(String xlpath1,String[][] xData) throws Exception
		{
			System.out.println("Inside XL Write");
			File myFile1=new File(xlpath1);
			FileOutputStream fout=new FileOutputStream(myFile1);
			HSSFWorkbook wb=new HSSFWorkbook();
			HSSFSheet mySheet1=wb.createSheet("TestResults");
			for(int i=0;i<xlRows;i++)
			{
				HSSFRow row1=mySheet1.createRow(i);
				for(short j=0;j<xlCols;j++)
				{
					HSSFCell cell1=row1.createCell(j);
					cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell1.setCellValue(xData[i][j]);
				}
			}
			wb.write(fout);
			fout.flush();
			fout.close();
		}
	}

