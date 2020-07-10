package test;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.*;

import report.Report;

public class SampleTest 
{
	@BeforeClass
	public static void setup() throws IOException
	{
		Report.initializeReport();
	}
	
	@Test(invocationCount = 5)
	public static void test1() throws AWTException, IOException
	{
		Report.updateReport("login", "login", "Pass");
	}
	@Test(invocationCount = 5)
	public static void test2() throws AWTException, IOException
	{
		Report.updateReport("Dashboard", "Dashboard", "Done");
	}
	@Test(invocationCount = 5)
	public static void test3() throws AWTException, IOException
	{
		Report.updateReport("logout", "logout", "Fail");
	}
	
	@AfterClass
	public static void teardown() throws InvalidFormatException, IOException
	{
		Report.consolidate_ScriptResult("Aditya Demo");
	}

}
