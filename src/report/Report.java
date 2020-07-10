package report;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import javax.imageio.ImageIO;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Report 
{
	public static Map<String, String> hmap = new HashMap<String, String>();
	public static String reportPath;
	public static String reportName;
	public static String reportFolderName;
	public static String strFinalPath;
	public static int stepNumber = 1;
	public static int strScreenFlag = 0;
	public static int steps = 1;
	public static int i = 0;
	public static int j = 0;
	public static int k = 0;
	public static final SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-YY hh-mm-ss-SSS");
	public static String timeStamp = sdf.format(new Date());
	public static LocalDateTime startTime;
	public static LocalDateTime endTime;
	
	
	public static String createResultDirectory()
	{
		String dir =  System.getProperty("user.dir")+"\\Result"+"\\Result-"+timeStamp;
		boolean createdDir= new File(dir).mkdir();
			if(createdDir)
			{
				System.out.println("Directory Created : "+dir);
			}
		return dir;
	}
	
	public static String TakeScreenshot(String Screenshot_flag) throws AWTException, IOException
	{
		String ScreenshotPath_Relative = "";
		if(Screenshot_flag.equalsIgnoreCase("Y"))
		{
			new File(createResultDirectory()+"/Screenshots").mkdir();
			String ssName = createResultDirectory()+"/Screenshots"+"/Step";
				   ssName = ssName + steps;
				   
			Robot robot = new Robot();
			Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
			BufferedImage screenFullImage = robot.createScreenCapture(screenRect);
			ImageIO.write(screenFullImage, "png", new File(ssName+ ".png"));
			
			ScreenshotPath_Relative = createResultDirectory()+"/Screenshots"+"/Step"+steps+".png";
			ssName = "";
			steps++;
			
		}
		return ScreenshotPath_Relative;	
	}
	
	public static void consolidateScreenshots() throws InvalidFormatException, IOException
	{
		@SuppressWarnings("resource")
		XWPFDocument doc =  new XWPFDocument();
		XWPFParagraph title = doc.createParagraph();
		XWPFRun run = title.createRun();
		
		title.setAlignment(ParagraphAlignment.CENTER);
		new File(createResultDirectory()+"/Consolidated Screenshots").mkdir();
		int count =  new File(createResultDirectory()+"/Screenshots").list().length;
		System.out.println("Total Number of Screen Captured :"+count);
		
		for (int i = 1; i < count; i++) 
		{
			run.addBreak();
			String imgFile = createResultDirectory()+"/Screenshots"+"/Step"+i+".png";
			FileInputStream fis = new FileInputStream(imgFile);
			run.addBreak();
			run.addPicture(fis, XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(500), Units.toEMU(400));
			fis.close();
		}
		FileOutputStream fos =  new FileOutputStream(createResultDirectory()+"/Consolidated Screenshots"+"/Consolidated Screenshots.docx");
		doc.write(fos);
		fos.close();
		System.out.println("Consolidated Screenshots.docx created");	
	}

	public static void consolidate_HTMLReport() throws IOException
	{
		endTime = LocalDateTime.now();
		long duration = ChronoUnit.SECONDS.between(startTime, endTime);
		reportPath = createResultDirectory()+"/HTML Reporting/Report.html";
		
		File htmlTemplateFile = new File(reportPath);
		String htmlString = FileUtils.readFileToString(htmlTemplateFile,StandardCharsets.UTF_8);
			   htmlString = htmlString.replace("$pCount", Integer.toString(i));
			   htmlString = htmlString.replace("$fCount", Integer.toString(k));
			   
			   if(duration >=3600) 
			   {
				   htmlString = htmlString.replace("$tcDurationHour", Long.toString(duration/3600));
				   htmlString = htmlString.replace("$tcDurationMins", Long.toString(duration/60 -60));
				   htmlString = htmlString.replace("$tcDurationSecs", Long.toString(duration%60));
			   }
			   else if(duration<3600)
			   {
				   htmlString = htmlString.replace("$tcDurationHour", "0");
				   htmlString = htmlString.replace("$tcDurationMins", Long.toString(duration/60 -60));
				   htmlString = htmlString.replace("$tcDurationSecs", Long.toString(duration%60));   
			   }
		File newHtmlFile = new File(reportPath);
		FileUtils.writeStringToFile(newHtmlFile, htmlString,StandardCharsets.UTF_8);
			   
	}
	
	public static void consolidate_HTMLReport(String TestName) throws IOException
	{
		endTime = LocalDateTime.now();
		long duration = ChronoUnit.SECONDS.between(startTime, endTime);
		InetAddress inetAddress = InetAddress.getLocalHost();
		reportPath = createResultDirectory()+"/HTML Reporting/Report.html";
		
		File htmlTemplateFile = new File(reportPath);
		String htmlString = FileUtils.readFileToString(htmlTemplateFile,StandardCharsets.UTF_8);
			   htmlString = htmlString.replace("$pCount", Integer.toString(i));
			   htmlString = htmlString.replace("$fCount", Integer.toString(k));
			   htmlString = htmlString.replace("$ScriptName", TestName);
			   htmlString = htmlString.replace("$OSName", System.getProperty("os.name"));
			   htmlString = htmlString.replace("$IPAddress", inetAddress.getHostAddress());
			   htmlString = htmlString.replace("$UserID", System.getProperty("user.name"));
			   
			   if(duration >=3600) 
			   {
				   htmlString = htmlString.replace("$tcDurationHour", Long.toString(duration/3600));
				   htmlString = htmlString.replace("$tcDurationMins", Long.toString(duration/60-60));
				   htmlString = htmlString.replace("$tcDurationSecs", Long.toString(duration%60));
			   }
			   else if(duration<3600)
			   {
				   htmlString = htmlString.replace("$tcDurationHour", "0");
				   htmlString = htmlString.replace("$tcDurationMins", Long.toString(duration/60));
				   htmlString = htmlString.replace("$tcDurationSecs", Long.toString(duration%60));   
			   }
		File newHtmlFile = new File(reportPath);
		FileUtils.writeStringToFile(newHtmlFile, htmlString,StandardCharsets.UTF_8);
			   
	}

	public static void consolidate_ScriptResult() throws InvalidFormatException, IOException
	{
		consolidateScreenshots();
		consolidate_HTMLReport();
	}
	
	public static void consolidate_ScriptResult(String TestName) throws InvalidFormatException, IOException
	{
		consolidateScreenshots();
		consolidate_HTMLReport(TestName);
	}
	
	public static void initializeReport() throws IOException
	{
		startTime = LocalDateTime.now();
		if(stepNumber <=1)
		{
			new File(createResultDirectory()+"/HTML Reporting").mkdir();
			
			File contentFile = new File(createResultDirectory()+"/HTML Reporting/content.jpg");	
			if(!contentFile.exists())
			{
				Files.copy(Paths.get(System.getProperty("user.dir")+"/src/report/content.jpg"), Paths.get(createResultDirectory()+"/HTML Reporting/content.jpg"));
			}
			File backgroundFile = new File(createResultDirectory()+"/HTML Reporting/background.jpg");	
			if(!backgroundFile.exists())
			{
				Files.copy(Paths.get(System.getProperty("user.dir")+"/src/report/background.jpg"), Paths.get(createResultDirectory()+"/HTML Reporting/background.jpg"));
			}
		
			reportPath = createResultDirectory()+"/HTML Reporting"+"/Report"+".html";
			File report_obj = new File(reportPath);
			
			if(report_obj.createNewFile())
			{
				System.out.println("File Created : "+report_obj.getName());
				
				String reportIn = new String(Files.readAllBytes(Paths.get(System.getProperty("user.dir")+"/src/report/resultformat.html")));
				Files.write(Paths.get(reportPath), reportIn.getBytes(), StandardOpenOption.CREATE);
		
				String report_content = new String(Files.readAllBytes(Paths.get(reportPath)));
				report_content.replaceFirst("ScriptDate$nbsp;ScriptTime", timeStamp);
				Files.write(Paths.get(reportPath), report_content.getBytes(), StandardOpenOption.CREATE);
			}
			else
			{
				System.out.println("File already present : "+report_obj.getName());
			}
		}
	}
	
	public static void updateReport(String StepName, String StepDesc, String Status) throws AWTException, IOException
	{
		String ScreenshotPath_Relative = TakeScreenshot("Y");
		String report_content = new String(Files.readAllBytes(Paths.get(reportPath)));
		
		if(Status.equalsIgnoreCase("Pass"))
		{
			i++;
			report_content = report_content.concat("<tr class='content'><td>"+ stepNumber +
												   "</td><td style='text-align:left;'>"+ StepName +"</td><td align='left'>"+ StepDesc +
												   "</td><td class='pass'>"+"<a style='color: #109015;font-weight: bold;' href="+"\""+ScreenshotPath_Relative+"\""+
												   "target=\"_self\">"+ Status +"</a>"+"</td><td>"+ sdf.format(new Date())+ "</td></tr>");
		}
		else if (Status.equalsIgnoreCase("Done"))
		{
			j++;
			report_content = report_content.concat("<tr class='content'><td>"+ stepNumber +
					   "</td><td style='text-align:left;'>"+ StepName +"</td><td align='left'>"+ StepDesc +
					   "</td><td class='done'>"+"<a style='color: #000000;font-weight: bold;'>"+ Status +"</a>"+"</td><td>"+ sdf.format(new Date())+ "</td></tr>");
		}
		else if (Status.equalsIgnoreCase("Fail"))
		{
			k++;
			report_content = report_content.concat("<tr class='content'><td>"+ stepNumber +
					   "</td><td style='text-align:left;'>"+ StepName +"</td><td align='left'>"+ StepDesc +
					   "</td><td class='fail'>"+"<a style='color: #ec0e0e;font-weight: bold;' href="+"\""+ScreenshotPath_Relative+"\""+
					   "target=\"_self\">"+ Status +"</a>"+"</td><td>"+ sdf.format(new Date())+ "</td></tr>");
		}
	
		Files.write(Paths.get(reportPath), report_content.getBytes(), StandardOpenOption.CREATE);
		stepNumber++;
	}
	
}
	
	
	

