// Copyright 2008 Vaau Inc. All Rights Reserved.
import java.io.*;

import jxl.*;
import jxl.read.biff.BiffException;

import java.text.*;

import java.io.*;
import java.util.Date;
import java.util.*;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

/**
 * This class takes file in Microsoft excel(xls) format and converts it to
 * feed file in csv format as per the schema specification.
 * @author deepak pandey (deepakpandey.21@gmail.com)
 *
 */
public class  ConvertCSV
{
    public static StringBuffer log = new StringBuffer();
    public static String jdbcDriver = new String();
    public static String jdbcUrl = new String();
    public static String jdbcUsername = new String();
    public static String jdbcPwd = new String();
    public static BufferedWriter xmlFile=null;

	public static Logger logger = null;
  public static LogManager manager = null;
  static {
    manager = LogManager.getLogManager();
    try {
      manager.readConfiguration(new FileInputStream("logging.properties"));
    } catch (SecurityException sE) {
      sE.printStackTrace();
    } catch (FileNotFoundException fnfE) {
      fnfE.printStackTrace();
    } catch (IOException ioE) {
      ioE.printStackTrace();
    }
    logger = Logger.getLogger(ConvertCSV.class.getName());
    manager.addLogger(logger);
    //logger.setLevel(Level.INFO);
    try {
      Date date = new Date();
      String day = date.getDate() + "." +date.getMonth() + "."+ date.getYear();
      logger.addHandler(new FileHandler("feedcsv" + day + ".log", true));
    } catch (SecurityException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
  /**
   * The main source of all evils.
   * @param args
   */
  public static void main(String[] args)
  {
    TreeSet output = new TreeSet();
    String xlsAcjrFiliFile = "";
    if (args.length != 1)
    {
      System.out.println("Usage: java ConvertCSV xlsInputFile");
      System.exit(1);
    }
    xlsAcjrFiliFile = args[0];

    try
    {
    	String encoding = "UTF8";
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("input.csv"), encoding));
       xmlFile = new BufferedWriter(new FileWriter("accpay_01_accounts.xml"));
         StringBuffer outXml = new StringBuffer();
      //Excel document to be imported
      WorkbookSettings ws = new WorkbookSettings();
      ws.setLocale(new Locale("en", "EN"));
      Workbook w = Workbook.getWorkbook(new File(xlsAcjrFiliFile),ws);
       outXml.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
            outXml.append("<rbacx>\n");
            outXml.append("<namespace namespaceName=\"" + "ACBS" + "\" namespaceShortName=\"" + "acbs" + "\" />\n<attributeValues>");
               outXml.append("</attributeValues>\n<accounts>");
      logger.log(Level.INFO, "Opening the sheets from XLS file in order...");
      // Gets the sheets from workbook
      for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++)
      {
        Sheet s = w.getSheet(sheet);

        Cell[] row = null;
        // Gets the cells from sheet
        for (int i = 0 ; i < s.getRows()-1 ; i++)
        {
          String roleName = null;
          String roleDescription = null;
          String endPoint = null;
          row = s.getRow(i);
          if (row.length >= 13)
          {
            roleName = row[0].getContents();
            roleDescription = row[1].getContents();
            endPoint = row[7].getContents();
          }
         // output.add("\"" + roleName  + "\",\"" + roleDescription + "\",\"" + endPoint + "\"");
            if(row[0].getContents().contains("PIN"))
            continue;

            outXml.append("<account id=\"");
                        outXml.append(row[0].getContents() );
                        outXml.append("\">\n");
                        outXml.append("<name><![CDATA[" + row[0].getContents().trim() + "]]></name>\n");
                        outXml.append("<endPoint>account Payable</endPoint>\n");
                        outXml.append("<domain></domain>\n");
                        outXml.append("<comments/>\n");
                        outXml.append("<attributes>\n");
                        outXml.append("<attribute name=\"Group ID\">\n");
                        outXml.append("<attributeValues>\n");
                        outXml.append("<attributeValue>\n");
                        outXml.append("<value>");
                        outXml.append("<![CDATA[" + row[1].getContents().trim() + "]]>");
                        outXml.append("</value>\n");
                        outXml.append("<attributes>\n");
                         outXml.append("<attribute name=\"Activity ID\">\n");
                        outXml.append("<attributeValues>\n");
                        outXml.append("<attributeValue>\n");
                        outXml.append("<value>");
                        outXml.append("<![CDATA[" + row[2].getContents().trim() + "]]>");
                        outXml.append("</value>\n");
                        outXml.append("<attributes>\n");
                        outXml.append("<attribute name=\"Activity Access\">\n");
                        outXml.append("<attributeValues>\n");
                        outXml.append("<attributeValue>\n");
                        outXml.append("<value>");
                        outXml.append("<![CDATA[" + row[4].getContents().trim()  + "]]>");
                        outXml.append("</value>\n");
                        outXml.append("<attributes>\n");

                        outXml.append("<attribute name=\"Flag\">\n");
                        outXml.append("<attributeValues>\n");
                        outXml.append("<attributeValue>\n");
                        outXml.append("<value>");
                        outXml.append("<![CDATA[" + row[5].getContents().trim()  + "]]>");
                        outXml.append("</value>\n");
                        outXml.append("</attributeValue>\n");
                        outXml.append("</attributeValues>\n");
                        outXml.append("</attribute>\n");
                        outXml.append("</attributes>\n");
                        outXml.append("</attributeValue>\n");
                        outXml.append("</attributeValues>\n");
                        outXml.append("</attribute>\n");
                        outXml.append("</attributes>\n");
                        outXml.append("</attributeValue>\n");
                        outXml.append("</attributeValues>\n");
                        outXml.append("</attribute>\n");
                        outXml.append("</attributes>\n");
                        outXml.append("</attributeValue>\n");
                        outXml.append("</attributeValues>\n");
                        outXml.append("</attribute>\n");
                        outXml.append("</account>\n");

        }
      }
      logger.log(Level.INFO, "The policy flag details.........");
      // TODO: Uncomment the following to get Console output.
      //ConvertCSV.displayFeed(output);
        outXml.append("</accounts>");
            outXml.append("</rbacx>");
      ConvertCSV.writeFeedToFile(outXml);
    }
    catch (IOException e)
    {
      logger.log(Level.SEVERE, e.toString() +
          " The XLS file format is not valid");
    }
    catch (BiffException be)
    {
      logger.log(Level.SEVERE, be.toString() +
          " The XLS file format is not valid");
    }
  }



public static void appendToLog(String businessUnitName, String message) throws Exception
  {
      DateFormat timeFormat = new SimpleDateFormat("hh:mm:ss");
      java.util.Date time = new java.util.Date();
      log.append("\n" + timeFormat.format(time) + " : ");

      if( message.equals("inactive") )
          log.append(" INACTIVE [" + businessUnitName + "]");
      else if( message.equals("add") )
          log.append(" ADDED    [" + businessUnitName + "]");
      else if( message.equals("update") )
          log.append(" UPDATED  [" + businessUnitName + "]");
      else if( message.equals("start") )
          log.append(" STARTED BUSINESS UNIT CREATION");
      else if( message.equals("finish") )
          log.append(" FINISHED BUSINESS UNIT CREATION");
		else
			log.append(message);
  }


  /**
   * Write the feed to the Console.
   * @param output
   */
  public static void displayFeed(Set output) {
	Iterator it = output.iterator();
    while (it.hasNext()) {
      String entry = (String)it.next();
      System.out.println(entry);
    }
  }

  /**
   * Write the feed to csv file.
   * @param output
   */
  public static void writeFeedToFile(StringBuffer output) {
     try {
    // FileOutputStream fos = new FileOutputStream("roles01");


      xmlFile.write(output.toString());
      xmlFile.close();
    } catch(IOException ioe) {
       ioe.printStackTrace();
    }
  }

}
