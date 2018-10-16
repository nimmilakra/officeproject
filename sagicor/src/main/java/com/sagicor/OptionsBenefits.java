package com.sagicor;
import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;

import com.relevantcodes.extentreports.ExtentTest;

public class OptionsBenefits
{

	static String ExpResultsFile = "";
	static String ActResultsFile = "D:\\Demo1\\ActualTest.xlsx";
	static String ExpSheetName = "ExpSheet";
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Nimmi30.txt";
	static String ActSheetName = "ActSheet";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\IUL01.pdf";
	

	public static void main(String[] args) throws Exception
	{
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Proposed Insured ", 15,1 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Initial Base Death Benefit", 15,2 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Initial Death Benefit Option", 15,3 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Target Premium", 15,4 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Annual Planned Premium Amount", 15,5 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Illustrated Premium Mode", 15,6 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "No Lapse Monthly Premium", 15,7 , ActSheetName);
		WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, "Life Insurance Qualification Test", 15,8 , ActSheetName);
		WriteExcelStrategy("Illustrated Account Allocations","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath,"Declared Rate Bonus Strategy", 15,9 , ActSheetName);
		WriteExcelStrategy("Illustrated Account Allocations","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath,"S&P 500® Index Bonus Strategy", 15,10 , ActSheetName);
		WriteExcelStrategy("Illustrated Account Allocations","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath,"Global Multi-Index Strategy", 15,11 , ActSheetName);
		//WriteExcelStrategy("ILLUSTRATION SUMMARY","INITIAL PREMIUM*",ExpResultsFile, ActResultsFile, TextFilepath, " ", 15,10 , ActSheetName);
		 
	}
	
	
	public static String WriteExcelStrategy(String FindValue, String TerminateValue, String ExpResultsFile,
			String ActResultsFile, String TextFilepath, String ColumnNAMe, int columnNumber, int rowNumberC,
			String ActSheetName) throws Exception
	{

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		try 
		{
			StringBuilder sb = new StringBuilder();
			
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 1;

			for (int i = 0; i <= lineNumber; i++) 
			{
				
				if (line.startsWith(FindValue))
				{
					
					for (int k = 0; k <= lineNumber; k++) 
					{
						if (line.startsWith(TerminateValue)) {
							break;
						}

						String restValue = line.replaceAll(ColumnNAMe, "");							
						restValue = restValue.trim();
							
							if(restValue.contains("%") && line.contains(ColumnNAMe) && Character.isDigit(restValue.charAt(0)))
							{
								PDFResults.setCellData(ActSheetName, columnNumber+1, rowNumberC, ColumnNAMe);						
								
							    PDFResults.setCellData(ActSheetName, columnNumber, rowNumberC, restValue);
							
							}
							else if(line.startsWith(ColumnNAMe))
							{
								PDFResults.setCellData(ActSheetName, columnNumber, rowNumberC, ColumnNAMe);
								PDFResults.setCellData(ActSheetName, columnNumber+1, rowNumberC, restValue);
								
							}
						
							
							line = br.readLine();
							lineNumber++;
						}
						
					}
				
				if (line.contains(TerminateValue)) {
					break;
				}
					line = br.readLine();
					lineNumber++;	
			}
			
			return "PASS";
			
			}
			catch (Exception e)
			{
			e.printStackTrace();
			return "FAIL";
			} 
			finally 
			{
			PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
			System.out.println("text to excel is Done");
			br.close();
			}
	}

		

	}


