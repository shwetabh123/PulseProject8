package generic;

import java.io.File;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.Test;

public class MoveFile {
	
	 protected static WebDriver driver;
	 public static String downloadPath = null;
	 
		@Test
	public  void MoveFileExample ()
	
	{
		
	
	    	try{
	    		
	    		
	    		
	    		BasePage b=new BasePage(driver);
	    		
	    		  downloadPath=b.preInitialize();
	    		
	    	   File afile =new File("C:\\Users\\Admin\\Downloads\\CompanyDirectoryParticipantImportTemplate_2.10.2019.xlsx");
	    		
	    	   if(afile.renameTo(new File("C:\\Users\\Admin\\workspace\\PulseProject8\\" + afile.getName()))){
	    		System.out.println("File is moved successful!");
	    	   }else{
	    		System.out.println("File is failed to move!");
	    	   }
	    	    
	    	}catch(Exception e){
	    		e.printStackTrace();
	    	}
	    }

}
