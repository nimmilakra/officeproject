package com.sagicor;
import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;

public class DataWithMultipleValuesClean 
{
	static String ExpResultsFile = "D:\\Demo1\\ExpectedSheetForData.xlsx";
	static String ActResultsFile = "D:\\Demo1\\ActualSheetForData.xlsx";
	static String ExpSheetName = "ExpSheet";
	//static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Vijay.txt";
	static String ActSheetName = "ActSheet";
	//static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC014.pdf";
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Nimmi3.txt";
	//static String pdfFilePath= "D:\\Sagicor_New_Final\\SEC014.pdf"
	

	public static void main(String[] args) throws Exception 
	{
		WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "Declared Rate Strategy", 0, 1, ActSheetName);
		WriteExcelStrategy(ExpResultsFile, ActResultsFile, TextFilepath, "S&P 500® Index Strategy", 0, 2, ActSheetName);

	}
	
	
	
	/**
     * Method to Compare the actual and expected values
     * @param testInst : for extent report instance.
     * @param ExpResultsFile : Expected file path.
     * @param ActResultsFile : Actual file path.
     * @param ExpSheetName : Expected sheet name.
     * @param ActSheetName : Actual sheet name.
     * 
     
     * @throws Exception If any exception occurred while compare the cells
     */
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
			int rowNumber = 2;

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
								
								if (Character.isDigit(splitDataSet[0].charAt(0)) && splitDataSet.length==4) 
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
