package com.pulse.Page;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import generic.BasePage;

public class HomePage extends BasePage {
	
	
	public HomePage(WebDriver driver) 
	
	
	{
		super(driver);
		PageFactory.initElements(driver,this);
	}
	
	
	
	BasePage b=new BasePage(driver);
	
	
	@FindBy(xpath="//*[@id='navbar-container']//ul/li/a[@href='#']")
	
	private WebElement  Arrowdownfor_settings ;
	
	
	

	@FindBy(xpath="//A[contains(., \"Logout\")]   ")
	
	private WebElement  Logout_click ;
	
	
	

	public void clickArrow(){
	
	//	Arrowdownfor_settings.click();
		
		b.Click(driver, Arrowdownfor_settings);
		
	}
	
	
	public void Logout(){
	
	//	Logout_click.click();
		
		b.Click(driver, Logout_click);
		
	}
	
	
	

}
