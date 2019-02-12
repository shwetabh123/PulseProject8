package com.pulse.Test;
	import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.By;
import org.testng.ITestResult;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.pulse.Page.Author;
import com.pulse.Page.HomePage;
import com.pulse.Page.LoginPage;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import generic.BasePage;
import generic.BaseTest;

import generic.Excel;
//import mx4j.log.Log;



//@Listeners(generic.RealGuru99TimeReport.class)


	public class ValidLogin3 extends BaseTest{

		public static   ExtentReports extent;
		public static  ExtentTest extentTest; 
		
		
		
	static int teststepcount=000000;
	
	
//	public  Logger APP_LOGS = Logger.getLogger("devpinoyLogger");

	
	  Excel eLib = new Excel();
	
	String url = eLib.getCellValue(pathexcel,"PreCon", 1, 0);
	
	  public static String logfiletimestamp;
	    
		  
	public static	ITestResult result = null;
		
	
	@Test
		public void testValidLogin3(Method method) throws Exception
		
		{
		


			
			driver.get(url);
			
	//		 Randomaplphanumber R=new Randomaplphanumber();
			  
	//		String r=  R.Random();
			
				String un=Excel.getCellValue(XLPATH,"ValidLogin",3,0);
				String pw=Excel.getCellValue(XLPATH,"ValidLogin",3,1);
				String accnt=Excel.getCellValue(XLPATH,"ValidLogin",3,2);
	
				String cb=Excel.getCellValue(XLPATH,"Author",1,3);
				
				
			
		
//		Logger APP_LOGS = Logger.getLogger(ValidLogin2.class);
				

				LoginPage l=new LoginPage(driver);
				

				
				BasePage b=new BasePage(driver);
				

				
				
	//			APP_LOGS.info("type username");
				
				
				//		l.setUserName(un);
						
						
						driver.findElement(By.xpath("//*[@id='j_username']")).sendKeys(un);;
						


//		        r= BaseTest.getScreenshot(driver, method.getName());

			
	//		APP_LOGS.info("type password");
				 
				 
				l.setPassword(pw);
				
				Thread.sleep(5000);

//		        r= BaseTest.getScreenshot(driver, method.getName());
					
					

			
	//			APP_LOGS.info("click select");
				
				l.clickLogin();
				
				Thread.sleep(5000);
				
				
	
				
	//	        r= BaseTest.getScreenshot(driver, method.getName());
					

	//			APP_LOGS.info("click dropdown ");
				
				
				
				l.dropdowntheaccount(accnt);
				Thread.sleep(5000);
			
				

					
	//		    r= BaseTest.getScreenshot(driver, method.getName());
						

	//			APP_LOGS.info("click select");
				
		
				l.clickselect();
				Thread.sleep(5000);
				

	//			r = BaseTest.getScreenshot(driver, method.getName());

			
				HomePage h=new HomePage(driver);
				
				h.clickArrow();
				

//				APP_LOGS.info("click arrrow down");
				
				h.Logout();
				
//			APP_LOGS.info("click logout");
		}

	


		}










