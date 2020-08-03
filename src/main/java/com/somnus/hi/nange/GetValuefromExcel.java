// 8.3 NinthTest
package com.somnus.hi.nange;

import java.io.FileInputStream;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.SearchContext;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.w3c.dom.Element;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public class GetValuefromExcel {
	public static String filestr;	
	static WebDriver driver;	
	
	// Basic data
	static int idname_row=2, idvalue_row=3;
	static int first_row=3;
	
    public static int SheetNum(String filestr) throws Exception
    {
    	InputStream ips=new FileInputStream(filestr); 
        XSSFWorkbook wb=new XSSFWorkbook(ips); 
        int num=wb.getNumberOfSheets();
        return num;
    }
	
	// Get column number according to column Title
	public static int col_num (int sn, String s) throws Exception 	{
		int i;
		for(i=1; !readCell(sn, 2,i).equals(s) && readCell(sn, 2, i)!=""; i++)
		{
			//System.out.println(readCell(2,i));
		}
		if(readCell(sn, 2, i)=="")
			return 0;
		else
			return i;			
	}
	
	// Get the number of column data according to column Title
	public static int data_num (int sn, String s) throws Exception {
		int i;
		int col=col_num(sn, s);
		for (i=3; readCell(sn, i, col).length()!=0; i++)
		{
			//System.out.println(readCell(i, col));
		}
		return i-3;			
	}
	
	public static int data_numNoExp(int sn, String s)
	{
		int x=0;
		try {
			x=data_num(sn, s);
		}
		catch(Exception ex)
		{}
		return x;
	}
	  	
	 // Get value from specified cell
    public static String readCell(int sheet_num, int row, int col) throws Exception{
    	InputStream ips=new FileInputStream(filestr); 
        XSSFWorkbook wb=new XSSFWorkbook(ips); 
        XSSFSheet sheet=wb.getSheetAt(sheet_num); 
        //int a= sheet.getPhysicalNumberOfRows();
        //System.out.println(a);
        
        if (row <= sheet.getPhysicalNumberOfRows())
	        {
		        XSSFCell cell=sheet.getRow(row-1).getCell(col-1);	
		        if(cell==null)
		        	return "";
		        else if(cell.getCellType()==XSSFCell.CELL_TYPE_NUMERIC)
		        	return cell.getStringCellValue();
		        else 
		        	return cell.getStringCellValue();   
	        }
        else
        	return "";
        } 
    
	public static String readCellNoExp(int sn, int a, int b)
	{
		String str = "";
		try
		{
			str=readCell(sn, a, b);
		}
		catch(Exception ex)
		{}
		return str;
	}
	
	public static WebElement NewWayFindElement(String s, String c)
	{
		switch(s) {
			case "ID":
				return driver.findElement(By.id(c));
			case "ClassName":
				return driver.findElement(By.className(c));
			case "Name":
				return driver.findElement(By.name(c));
			default:
				return null;	
		}
	}
	
	public static WebElement NewWayFindElement(String s, String c, String r)
	{
		switch(s) {
			case "ID&ClassName":
				return driver.findElement(By.id(c)).findElement(By.className(r));
			default:
				return null;	
		}
	}
     
	// Preoperation
	public static void Preoper(String op) throws Exception
	{
		switch(op) {
			case "a":
				System.out.println("There is a preoperation \"a \" ");
				driver.findElements(By.id("unit_type_chosen")).get(1).findElement(By.className("chosen-search-input")).sendKeys("centi - 0.01"+"\n");   
		        break;
			case "b":
				System.out.println("There is a preoperation \"b \" ");
				driver.findElements(By.id("unit_type_chosen")).get(1).findElement(By.className("chosen-search-input")).sendKeys("exa - 1.0E18"+"\n");   
		        break;
			case "c":
				System.out.println("There is a preoperation \"c \" ");
		        driver.get("http://dssat2d-plot.herokuapp.com/demo/vmapper");
		       /* Thread.sleep(1000);
		        driver.manage().window().maximize();
		        Thread.sleep(1000);
		        driver.getCurrentUrl();
		        driver.findElement(By.xpath("/html/body/div[12]/div/div/div[2]/div/div/div[1]/div/span/label/span")).click();
		        Thread.sleep(1000);
		        driver.getCurrentUrl();
		        Thread.sleep(1000);
		        driver.getCurrentUrl()
		        driver.findElement(By.xpath("/html/body/div[12]/div/div/div[2]/div/div/div[1]/div/input")).sendKeys("C:\\Users\\ga_ri\\Desktop\\Preop\\cng201819_2020-05-21_jgt_chp.xlsx");
		        Thread.sleep(1000);
		        System.out.println("hello!");
		        */

		        //.sendKeys("C:\\Users\\ga_ri\\Desktop\\Preop\\cng201819_2020-05-21_jgt_chp.xlsx");
		       // driver.findElement(By.xpath("/html/body/div[12]/div/div/div[2]/div/div/div[2]/div/input")).sendKeys("C:\\Users\\ga_ri\\Desktop\\Preop\\cng201819_2020-05-21_jgt_chp-sc2.json");
		       // driver.findElement(By.xpath("/html/body/div[12]/div/div/div[3]/button[2]")).click();
		        //driver.quit();
		        driver.get("http://dssat2d-plot.herokuapp.com/demo/unit");
		        driver.manage().window().maximize();
			case "":
				break;
	        default:
	        	
		}
	}
	
    // Test first module
    public static void Test(WebDriver driver, int sheet_num) throws Exception
    {
    	// int sheet_num=1;
    	
    	    System.out.println("/******************************************************************************************/");
    	    System.out.println("Following is the \" " + readCell(sheet_num, 1, 1) + " \" module test:" + "\n");

	        for(int i=first_row ; readCell(sheet_num, i , col_num(sheet_num, "Input_Cases_1"))!=""; i++)
	    	{
	        	// preoperation
	        	if(col_num(sheet_num, "Preoperation")!=0)
	        	{
	        		Preoper(readCell(sheet_num, i, col_num(sheet_num, "Preoperation")));
	        	}	        	
	        	
	        	// input all the input datum
	        	if(col_num(sheet_num, "Input_ClassName")!=0)
	        	{
		        	for(int a=1; a<=data_num(sheet_num, "Input_ID"); a++)
		        	{
		        		NewWayFindElement(readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_Element_Find")),readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_Element_Find")+1),readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_Element_Find")+2)).sendKeys(readCell(sheet_num, i,col_num(sheet_num, "Input_Cases_1")+a-1)+"\n");
		        		try 
		     	        { 
		     	        	Thread.sleep (1000) ;
		     	        } 
		     	        catch (InterruptedException ie){}	     	       
		        	}
	        	}
	        	else
	        	{
		        	for(int a=1; a<=data_num(sheet_num, "Input_ID"); a++)
		        	{
		        		NewWayFindElement(readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_Element_Find")),readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_Element_Find")+1)).sendKeys(readCell(sheet_num, i,col_num(sheet_num, "Input_Cases_1")+a-1));		        		
		        	}
	        	}

	        	try 
	 	        { 
	 	        	Thread.sleep (1000) ;
	 	        } 
	 	        catch (InterruptedException ie){}
	        	
	        	// compare the results
	        	for(int a=1; a<=data_num(sheet_num, "Output_ID"); a++)
	        	{
	        		String expect_str=readCell(sheet_num, i, col_num(sheet_num, "Expect_result_1")+a-1);
	        		String actual_str=driver.findElement(By.id(readCell(sheet_num, idvalue_row+a-1 , col_num(sheet_num, "Output_ID")))).getText();
	        		if(expect_str.equals(actual_str))
	        			System.out.println(i-first_row+1 + " Test case (" + readCell(sheet_num, i,col_num(sheet_num, "Input_Cases_1")) + ") test \"" + readCell(sheet_num, first_row-1, col_num(sheet_num, "Expect_result_1")+a-1) + "\" passed.");
				    else
				    	System.out.println(i-first_row+1 + " Test case (" + readCell(sheet_num, i,col_num(sheet_num, "Input_Cases_1")) + ") test \"" + readCell(sheet_num, first_row-1, col_num(sheet_num, "Expect_result_1")+a-1) + "\" false.");
	        	}		      

	        	// clear the data for the next test data
	        	for(int a=1; a<=data_num(sheet_num, "Input_ID"); a++)
	        	{
	        		if(col_num(sheet_num, "Input_ClassName")!=0)
	        		{}
	        		else
	        		{
	        			driver.findElement(By.id(readCell(sheet_num, idvalue_row+a-1, col_num(sheet_num, "Input_ID")))).clear();
	        		}	        			        			
	        	}
			 }
	        System.out.println("\"" + readCell(sheet_num, 1,1) +"\" module test ends. \n");
	    }        
    }
