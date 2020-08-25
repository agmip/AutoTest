// 8.23 NinthTest, preoperation

package com.somnus.hi.nange;
import java.awt.List;
import javax.lang.model.element.Element;
import javax.swing.text.html.HTMLDocument.Iterator;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import com.google.common.collect.Lists;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.*;
import java.io.*;
import java.net.BindException;
import java.util.ArrayList;

public class App { 
 
	public static void main(String[] args) throws Exception { 
    	GetValuefromExcel abc=new GetValuefromExcel();
    	
    	// A new testing work, you need to change the testing folder.
    	String testFolder="C://Users//ga_ri//Desktop//TestFolder";
    	
    	int filenumber=GetValueformFile.GetFileNum(testFolder);
    	File filestr[];
    	filestr=GetValueformFile.GetAllFiles(testFolder);
    	System.out.println("There are " + filenumber + " test files in the folder.");
    	
    	System.setProperty("webdriver.chrome.driver","C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe");
    	// declaration and instantiation of objects/variables
        WebDriver driver = (WebDriver) new ChromeDriver();
        GetValuefromExcel.driver=driver;
              
    	// A new testing work, you need to change the website.
        // Launch website
        driver.get("http://dssat2d-plot.herokuapp.com/demo/unit");
        driver.manage().window().maximize();
        String titile = driver.getTitle();
        
        System.out.println("Test title is => " + titile + "\n");
 
        for(int f=0; f<GetValueformFile.GetFileNum(testFolder); f++)
        {
        	abc.filestr=filestr[f].toString();
        	System.out.println("/******************************************************************************************m/");
        	System.out.println("/******************************************************************************************m/");
    	    System.out.println("Following is the tests in file: \" " + abc.filestr + "\n");
        	int num=abc.SheetNum(abc.filestr);
        	
            for(int i=0; i<num; i++)
            {
            	abc.Test(driver, i);
            }                   	
        }

    	driver.close();     
        System.out.println("The testing has been completed!");
    }
    
}
