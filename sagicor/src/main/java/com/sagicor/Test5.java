package com.sagicor;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;

public class Test5 
{

	static String ExpResultsFile = "";
	static String ActResultsFile = "D:\\demo1\\ActualSheet.xlsx";
	static String ExpSheetName = "ExpSheet";
	static String TextFilepath= "D:\\Sagicorpdf\\PDFToText_pdf1.txt";
	static String ActSheetName = "ActSheet";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\IUL01.pdf";
    
	public static void main(String[] args) throws Exception
	{
		//WriteExcelStrategy("","",ExpResultsFile, ActResultsFile, TextFilepath, "Proposed Insured ", 15,1 , ActSheetName);
		WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "December", 0, 1, ActSheetName);

	}

	public static String WriteExcelStrategy(String ExpResultsFile,
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
						if (line.startsWith(ColumnNAMe))
						{
							//PDFResults.setCellData(ActSheetName, columnNumber, rowNumberC, ColumnNAMe);
							//String restValue = line.replaceAll(ColumnNAMe, "");
							String restValue=line.replaceAll(" %","%");
							//restValue = restValue.trim();
							String[] splitDataSet = restValue.split("\\s+");
							String str=splitDataSet[0]+splitDataSet[1]+splitDataSet[2];
							PDFResults.setCellData(ActSheetName, columnNumber, rowNumberC,str);
							
                             
							for (int j = 3; j < splitDataSet.length; j++)
							{
								System.out.println(splitDataSet.length);
								
								if (Character.isDigit(splitDataSet[3].charAt(0)) && splitDataSet.length==11) 
								{
									String data1 = splitDataSet[j];
									
									setCellList_intColumn.add(j-2);
									setCellList_intRow.add(rowNumberC);
									setCellList_Str.add(data1);
									//System.out.println(splitDataSet.length);
									
								}
								if (splitDataSet.length == j+1) 
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
