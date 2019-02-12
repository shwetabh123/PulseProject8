package com.pulse.Test;


	import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.testng.ITestResult;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.pulse.Page.Author;
import com.pulse.Page.CompanyDirectory;
import com.pulse.Page.HomePage;
import com.pulse.Page.LoginPage;
import com.relevantcodes.extentreports.LogStatus;

import generic.BasePage;
import generic.BaseTest;

import generic.Excel;
//import mx4j.log.Log;



//@Listeners(generic.RealGuru99TimeReport.class)


	public class AddParticipant extends BaseTest{

	
	static int teststepcount=000000;
	
	
	public static Logger APP_LOGS = Logger.getLogger("devpinoyLogger");
	  Excel eLib = new Excel();
		
	String url = eLib.getCellValue(pathexcel,"PreCon", 1, 0);
	
	  public static String logfiletimestamp;

		public static Properties CONFIG;
	  
	
	  
	//public static	ITestResult result = null;
		
		@Test(priority=1)
		public void addparticipant(Method method) throws Exception
		
		{
			
			
			
			driver.get(url);
			
			
			Thread.sleep(15000);
			
			
	//		 Randomaplphanumber R=new Randomaplphanumber();
			  
	//		String r=  R.Random();
			
				String un=Excel.getCellValue(XLPATH,"ValidLogin",3,0);
				String pw=Excel.getCellValue(XLPATH,"ValidLogin",3,1);
				String accnt=Excel.getCellValue(XLPATH,"ValidLogin",3,2);
	
				String cb=Excel.getCellValue(XLPATH,"Author",1,3);
			       
				//Logger APP_LOGS = Logger.getLogger("devpinoyLogger");
			     
				Logger APP_LOGS = Logger.getLogger(AddParticipant.class);
				
	    		    
				//String HPT=Excel.getCellValue(XLPATH,"ValidLogin",1,2);
				CompanyDirectory cd=new CompanyDirectory(driver);
				
				LoginPage l=new LoginPage(driver);
				
		
				
				l.setUserName(un);
				Thread.sleep(15000);
				
				APP_LOGS.info("type username");
				
				l.setPassword(pw);
				
				Thread.sleep(15000);
				APP_LOGS.info("type password");
				
				l.clickLogin();
				
				Thread.sleep(15000);
				
				APP_LOGS.info("click login");
				l.dropdowntheaccount(accnt);
				
				Thread.sleep(15000);
				
				APP_LOGS.info("click dropdown ");
				
				
				l.clickselect();
				
				Thread.sleep(15000);
				
			
				APP_LOGS.info("click select");
			//	cd.clickArrow();
				
				Thread.sleep(15000);
				
				
				
				cd.clickSettings();
				
				Thread.sleep(15000);
				
				
				APP_LOGS.info("click setings");
				
				cd.clickCompanyDirectory();
				Thread.sleep(15000);
				
				APP_LOGS.info("click CD");
				
				cd.clickParticipants();
				Thread.sleep(15000);
				
				
				
				cd.clickuploadparticipantscompany();
				
				Thread.sleep(15000);
				
				
			
				
			BasePage b=new BasePage(driver);
			
			String data=b.preInitialize();
			
			
				b.getDownloadedFileDetails(data);
				Thread.sleep(15000);
				
				
				
				
				cd.clickDownloadSampleImportTemplate();
				
				Thread.sleep(15000);
				
				
				
			//	b.getLastDownloadedFile("GetLastDownloadedFile");
				
				
				String fileName =b.getLastDownloadedFile("GetLastDownloadedFile");
				
			
		    		
				
				System.out.println("Last downloaded file is "+fileName);
				
				
			
			
				
				Thread.sleep(15000);
				
				
		//		b.addeditTextinExcel(downloadPath, fileName, "Employee File", 0, 2, "charu2");
	
		    	
				  
		}











}
