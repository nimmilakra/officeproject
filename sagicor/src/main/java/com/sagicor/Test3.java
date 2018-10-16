package com.sagicor;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;


public class Test3 
{
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Nimmi30.txt";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\IUL01.pdf";
	static String ActResultsFile = "D:\\Demo1\\ActualNewTestDemo.xlsx";
	static String ActSheetName = "ActSheet";

	public static void main(String[] args) throws Exception
	{
		Output_HiLowSPFIA14_ReadExcel(ActResultsFile,TextFilepath,
				ActSheetName, pdfFilePath);		

	}
	
	public static String Output_HiLowSPFIA14_ReadExcel(String ActResultsFile, String TextFilepath, String ActSheetName, String pdfFilePath) throws Exception 
	{
		
		Xlsx_Reader PDFResults = new Xlsx_Reader(ActResultsFile);
		BufferedReader br = new BufferedReader(new FileReader(TextFilepath));
		ArrayList<Integer> setCellList_intColumn = new ArrayList<Integer>();
		ArrayList<Integer> setCellList_intRow = new ArrayList<Integer>();
		ArrayList<String> setCellList_Str = new ArrayList<String>();
		try 
		{
			StringBuilder sb = new StringBuilder();
			// String line1 = br.readLine();
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 1;
			
			 for (int i = 0; i <= lineNumber; i++) 
			{
						if(line.contains("December"))
						{
							//line.replace(" 31","31");
							//String[] splitDataSet = line.split("\\s+");
							String[] splitDataSet=line.replaceAll(" %","%").split("\\s+");
							
							
							 System.out.println("splitData Length=" + splitDataSet.length);
							 String data1 = splitDataSet[0]+ splitDataSet[1]+splitDataSet[2];
							   setCellList_intColumn.add(0);	
								setCellList_intRow.add(rowNumber);
							    setCellList_Str.add(data1);
							for (int j =3; j < splitDataSet.length; j++) 
							{
								  
									if((splitDataSet[0].contains("December") &&  (splitDataSet.length==11)))		
									{
										
										String data2=splitDataSet[j];
										setCellList_intColumn.add(j);	
										setCellList_intRow.add(rowNumber);
									    setCellList_Str.add(data2);
									if (splitDataSet.length == j + 1) 
									{
										rowNumber++;
									}
									
									
								}
							}
						}
						
						
							sb.append(line);
							sb.append(System.lineSeparator());
						
						
						line = br.readLine();
						lineNumber++;
					
					
						
						
				
			}
			
			String everything = sb.toString();
			// System.out.println(everything);
			
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


