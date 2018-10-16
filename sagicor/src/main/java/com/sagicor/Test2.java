package com.sagicor;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;

public class Test2 
{
	

	static String ExpResultsFile = "D:\\Demo1\\ExpectedDemoNewTable.xlsx";
	static String ActResultsFile = "D:\\Demo1\\ActualDemoNewTable.xlsx";
	static String ExpSheetName = "ExpSheet";
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Nimmi3.txt";
	static String ActSheetName = "ActSheet";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC017.pdf";


	public static void main(String[] args) throws Exception
	{
		
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Prepared for", 0, 1, ActSheetName);
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Prepared by", 0, 2, ActSheetName);
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Prepared on", 0, 3, ActSheetName);
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Premium Payment", 0, 4, ActSheetName);
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Tax Qualification", 0, 5, ActSheetName);
		 WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Issue State", 0, 6, ActSheetName);

	}
	
	public static String WriteExcelStrategy( String ExpResultsFile,
			String ActResultsFile, String TextFilepath, String ColumnNAMe, int columnNumber, int rowNumberC,
			String ActSheetName) throws Exception {

		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		try {
			StringBuilder sb = new StringBuilder();
			String line = br.readLine();
		
			int lineNumber = 1;
			int rowNumber = 1;

			for (int i = 0; i <= lineNumber; i++) 
			{
						if (line.contains(ColumnNAMe))
						{
							
							PDFResults.setCellData(ActSheetName, columnNumber, rowNumberC, ColumnNAMe);
							
						
							String restValue = line.replaceAll(ColumnNAMe, "");
							restValue = restValue.trim();
							String[] splitDataSet = restValue.split("\\s+");
							
                             
							for (int j = 0; j < splitDataSet.length; j++)
							{
								
								if (Character.isDigit(splitDataSet[0].charAt(0))  || Character.isLetter(splitDataSet[0].charAt(0))|| splitDataSet[0].contains("$")  ) 
								{
									String data1 = splitDataSet[j];
									setCellList_intColumn.add(j+1);
									setCellList_intRow.add(rowNumberC);
									setCellList_Str.add(data1);
									
								}
								if (splitDataSet.length == j+2) 
								{
									rowNumber++;
								}
							}
							
							sb.append(line);
							sb.append(System.lineSeparator());
						}
						line = br.readLine();
						lineNumber++;
					}
				
					line = br.readLine();
					lineNumber++;
				
				String everything = sb.toString();
				
			
			return "PASS";

		} catch (Exception e) {
			e.printStackTrace();
			return "FAIL";
		} finally {
			PDFResults.setCellData_Perform(ActSheetName, setCellList_intColumn, setCellList_intRow, setCellList_Str);
			System.out.println("text to excel is Done");
			br.close();
		}

	}
	

}


