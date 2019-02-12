package generic;



import org.testng.ITestResult;
import org.testng.annotations.AfterClass;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.RollingFileAppender;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Platform;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxBinary;
import org.openqa.selenium.firefox.FirefoxDriver;
//import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.logging.LogType;
import org.openqa.selenium.logging.LoggingPreferences;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;

import com.pulse.Page.HomePage;
import com.pulse.Page.LoginPage;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;




public abstract class BaseTest implements IAutoConst{
	

	public    RemoteWebDriver driver=null;

	public static int REPEAT_MINUS = 100;
	
Logger APP_LOGS = Logger.getLogger("BaseTest");
     
	
	public static String scrshotFolderLoc;


	  // public static String pathexcel ="C:\\Users\\Admin\\workspace\\PulseProject8\\data\\preCondInput.xlsx";

	   final static String workingDirexcel = System.getProperty("user.dir");
		
		  //  final static String filePath = "./test-output/MyReport.html";
		  final static String filePathexcel = "\\data\\preCondInput.xlsx";

		  public static String pathexcel =workingDirexcel+filePathexcel;
		
		
		
		 final static String workingDir = System.getProperty("user.dir");
			
		  //  final static String filePath = "./test-output/MyReport.html";
		  final static String filePath = "\\test-output\\MyReport.html";

		  public static String path =workingDir+filePath;
		
			
	

	
	static int testRunId=000000;
	
	  public static String logfiletimestamp;

		 
		 public static String downloadPath = null;
		
	//	public  ExtentReports extent;
	//	public  ExtentTest extentTest; 
		
	
	    
	    
		static  Excel eLib = new Excel();

		
	//	String browserType = eLib.getCellValue(path,"PreCon", 1, 1);
		
		
	//	 int time = Integer.parseInt((eLib.getCellValue(path,"PreCon", 1, 2)).replace(".0", ""));
		
		/*
		@BeforeTest
		public void setExtent()
		
		
		{
			extent = new ExtentReports(System.getProperty("user.dir")+"/test-output/"+"AutomationLog_TestRunId ExtentReport.html", false);
			extent.addSystemInfo("Host Name", "Naveen Mac");
			extent.addSystemInfo("User Name", "Naveen Automation Labs");
			extent.addSystemInfo("Environment", "QA");
			
			
			
		}*/
		
		@BeforeClass
		
		@Parameters({"browser"})
		
		public void setUp(String browser) throws Exception 
		
		{
			
		     
			
			
		/*	extent = new ExtentReports(System.getProperty("user.dir")+"/test-output/"+"AutomationLog_TestRunId ExtentReport.html", false);
			
			extentTest = extent.startTest("testName", "description");
			
			
			extent.addSystemInfo("Host Name", "Naveen Mac");
			extent.addSystemInfo("User Name", "Naveen Automation Labs");
			extent.addSystemInfo("Environment", "QA");
			extent.startTest("Test Name","Description");*/
			
			
			System.out.println("*******************");
			
			
				
				
			driver = getDriver(browser);
			
		
			driver.manage().window().maximize();
			
	
		}
		

		
		


		
		public static  RemoteWebDriver getDriver(String browser) throws Exception {
	
			//extentTest =extent.startTest(method.getName());
			
			 String fileName=null;
			   
	    	 Date d = new Date();
	    	 
	    	 
	 	    SimpleDateFormat format = new SimpleDateFormat(  "dd.MMMMM.yyyy.hh.mm");
	 	   

	 	   logfiletimestamp = "AutomationLog_TestRunId_"+format.format(d);
	 	   
		    System.setProperty("autologname",logfiletimestamp);

		      PropertyConfigurator.configure("Log4j.properties");
		    
        	String browserType = eLib.getCellValue(path,"PreCon", 1, 1);
	

		
		//	return new RemoteWebDriver(new URL("http://10.111.108.157:4444/wd/hub"), getBrowserCapabilities(browser));
        	return new RemoteWebDriver(new URL("http://192.168.2.6:4444/wd/hub"), getBrowserCapabilities(browser));
			
		}
		
		public static   DesiredCapabilities getBrowserCapabilities(String browser) throws Exception
		
		{
			
		    
		
			 
            DesiredCapabilities dr=null;

            
            RemoteWebDriver driver=null;
            
			switch (browser) 
			
			
			
			{
			case "firefox":
			
			System.out.println("Opening firefox driver");

		//	System.setProperty("webdriver.chrome.driver", "C:\\Users\\ssrivastava4\\workspace\\PulseProject7\\driver\\chromedriver.exe");
		
			

			 DesiredCapabilities capability = new DesiredCapabilities().firefox();
	            capability.setBrowserName("firefox");
	            capability.setPlatform(Platform.WINDOWS);
	            
	         //   DesiredCapabilities capabillities = new DesiredCapabilities("firefox", "3.6.", Platform.WINDOWS);
	            capability.setCapability("job-name", "Fancy Firefox profile");
	            FirefoxProfile profile = new FirefoxProfile();
	            profile.setPreference("network.http.phishy-userpass-length", 255);
	            capability.setCapability(FirefoxDriver.PROFILE, profile);
	            
	            
	            
	            
			return DesiredCapabilities.firefox();
	            		
				
			case "chrome":
				System.out.println("Opening chrome driver");
			
			
				    DesiredCapabilities capability1 = DesiredCapabilities.chrome();
				    capability1.setBrowserName("chrome");
				    capability1.setPlatform(Platform.ANY);


				    ChromeOptions options = new ChromeOptions();
		          //  options.addArguments("--start-maximized");

				    return DesiredCapabilities.chrome();
					

				//	APP_LOGS.info("Chrome Browser opened");
				
				
				
			case "IE":
				System.out.println("Opening IE driver");
				return DesiredCapabilities.internetExplorer();
			default:
				System.out.println("browser : " + browser + " is invalid, Launching Firefox as browser of choice..");
				return DesiredCapabilities.firefox();
				
				
			
			
			}
					
		}
		
		
	
		
		
		
	

		
		
		public static String getScreenshot(WebDriver driver, String screenshotName) throws IOException{
			
			
			String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
			TakesScreenshot ts = (TakesScreenshot) driver;
			File source = ts.getScreenshotAs(OutputType.FILE);
			// after execution, you could see a folder "FailedTestsScreenshots"
			// under src folder
			String destination = System.getProperty("user.dir") + "/screenshots/" + screenshotName + dateName
					+ ".png";
			File finalDestination = new File(destination);
			FileUtils.copyFile(source, finalDestination);
			return destination;
		}

	
		
	
/*	
	@AfterMethod

	
		
		public void tearDown(ITestResult result) throws IOException
	{
			
			
		
		
		
			if(result.getStatus()==ITestResult.FAILURE){
				extentTest.log(LogStatus.FAIL, "TEST CASE FAILED IS ->"+result.getName()); //to add name in extent report
				extentTest.log(LogStatus.FAIL, "TEST CASE FAILED IS ->"+result.getThrowable()); //to add error/exception in extent report
				
				String screenshotPath = BaseTest.getScreenshot(driver, result.getName());
				
				
				
				extentTest.log(LogStatus.FAIL, extentTest.addScreenCapture(screenshotPath)); //to add screenshot in extent report
			
				
				//extentTest.log(LogStatus.FAIL, extentTest.addScreencast(screenshotPath)); //to add screencast/video in extent report
		
			
			}
			else if(result.getStatus()==ITestResult.SKIP){
				extentTest.log(LogStatus.SKIP, "Test Case SKIPPED IS " + result.getName());
			}
			
			
			else if(result.getStatus()==ITestResult.SUCCESS){
				
				extentTest.log(LogStatus.PASS, "TEST CASE PASSED IS ->"+result.getName()); //to add name in extent report
				
				
				//extentTest.log(LogStatus.PASS, "TEST CASE PASSED IS ->"+result.getThrowable()); //to add error/exception in extent report
				
			//	String screenshotPath1 = BaseTest.getScreenshot(driver, result.getName());
				
				
			//	extentTest.log(LogStatus.PASS, extentTest.addScreenCapture(screenshotPath1));
		

			}
			
			HomePage hp = new HomePage(driver);
			//hp.clickOnProfile(driver);
			hp.clickArrow();
			
			hp.Logout();
	
		
			
			
			extent.endTest(extentTest); //ending test and ends the current test and prepare to create html report
			
			
	}
	
	
	*/
	
	
	
/*
	
	@AfterTest
	public void postCondition() throws InterruptedException
	
	
	{
		
		//driver.get("D:\\ShwetabhWorkspace\\PulseProject8\\test-output\\"+"AutomationLog_TestRunId-"+b.getlastRunid()+"ExtentReport.html");
		
		
		
		driver.get("D:\\ShwetabhWorkspace\\PulseProject8\\test-output\\"+"AutomationLog_TestRunId-ExtentReport.html");
		
		
		
		
	//	Thread.sleep(25000);
		
		//driver.quit();
		extent.flush();
		//extent.close();
		}
*/	
	
	
	
	   @AfterClass
	    public void tearDown() throws Exception {
	        driver.quit();
	    }
			

				
		
		
		
	

		
		
			
}	





