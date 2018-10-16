package com.sagicor;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class NewDemo 
{
	static String TextFilepath= "D:\\Sagicor_New_Final\\PDFToText_Nimmi30.txt";
	static String pdfFilePath= "D:\\Sagicor_New_Final\\IUL01.pdf";
	static String ActResultsFile = "D:\\Demo1\\ActualNewTestDemo.xlsx";
	static String ActSheetName = "ActSheet";

	public static void main(String[] args) throws Exception, Exception
	{
		pdftoText(pdfFilePath, TextFilepath);
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
		
			String line = br.readLine();
			int lineNumber = 1;
			int rowNumber = 1;
			
			outerloop: for (int i = 0; i <= lineNumber; i++) 
			{
		      
						if (Character.isDigit(line.charAt(0)))
						{
							String[] splitDataSet = line.split("\\s+");
							
                    
							for (int j = 0; j < splitDataSet.length; j++) 
							{
								if( (splitDataSet[0].length()==1||splitDataSet[0].length()==2)  && (splitDataSet.length==13))
										
								{
									String data1 = splitDataSet[j];
									// setCellData(String sheetName,int colName,int rowNum, String data)
									setCellList_intColumn.add(j);	
									setCellList_intRow.add(rowNumber);
									setCellList_Str.add(data1);
									//PDFResults.setCellData(ActSheetName, j + 21, rowNumber, data1);
								

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
	
	
	public static String pdftoText(String pdfFilePath, String TextFilepath) throws InterruptedException, IOException
	{
		// APP_LOGS.debug("Click on Button");
		try 
		{
			PDFParser parser;
			String parsedText = null;
			PDFTextStripper pdfStripper = null;
			PDDocument pdDoc = null;
			COSDocument cosDoc = null;

			pdDoc = PDDocument.load(new File(pdfFilePath));
			pdfStripper = new PDFTextStripper();
			String content = pdfStripper.getText(pdDoc);

			File file = new File(TextFilepath);

			// if file doesnt exists, then create it
			if (!file.exists()) {
				file.createNewFile();
			}

			FileWriter fw = new FileWriter(file.getAbsoluteFile());
			BufferedWriter bw = new BufferedWriter(fw);
			bw.write(content);
			bw.close();

			System.out.println("Pdf to text Done");
			return "PASS";

		} catch (IOException e) {
			e.printStackTrace();
			return "FAIL";
		}
	}
	

}

