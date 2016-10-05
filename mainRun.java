/* Developer: John Albert Herrick II
 * Created by Developer on 10/4/16.
 * Copyright © 2016 John Albert Herrick II. All rights reserved.
 * 
 * In affiliation with the Center of Teaching Excellence at the University of Virginia
 * This is a general .xls reader for purposes of consolidation of observation data of professors.
 * Input requirements: files to be read must be in Excel 97-2004.xls format to work.
 * See the README file for use instruction and more details.
 * 
 */


package lib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormatter;

public class mainRun {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		  //get all files into a list
		  ArrayList<String> files = new ArrayList<String>();
		  File dir = new File("/Users/johnherrick/Documents/Eclipse/Teaching Excellence");
		  File[] directoryListing = dir.listFiles();
		  if (directoryListing != null) {
		    for (File child : directoryListing) {
		    	if(child.toString().contains("xls") && !(child.toString().contains("Coded Data")) && !(child.toString().contains("output1")))
		    		files.add(child.toString().substring(57)); 	    
		    }    
		  }
		  
		  
		
		//Set new workbook (The Output)
		HSSFWorkbook writenew = new HSSFWorkbook();
		HSSFSheet output = writenew.createSheet("output");
		
		
		
		//Set the initial headings of new workbook
		String[] headings = {"Instructor ID", "Instructor", "ClassSubject",	"Treatment", "Program",	"ProgramYear", "ClassNumber", 
				"ClassSize", "ClassLevel", "ClassDate",	"TotalTime", "Observation",	"Observer",	"ObserverID",  "I-1o1", "I-Adm",	
				"I-AnQ", "I-CQ", "I-DV", "I-FUp", "I-Lec", "I-MG", "I-Other", "I-PQ", "I-RtW", "I-W", "I-_Si", "Notes", "S-AnQ", "S-CG", 
				"S-Ind", "S-L", "S-OG", "S-Other", "S-Prd", "S-SP", "S-SQ",	"S-TQ", "S-W", "S-WC", "S-WG"};
		int rowCount = 0;
		HSSFRow headRow = output.createRow(rowCount);
		for(int i =0; i<headings.length; i++){
			HSSFCell cell = headRow.createCell(i);
			cell.setCellValue(headings[i]);
		}
		rowCount++;
		
		
		
		//Loop through each file and perform all retrieval and averaging methods below. Add new rows to the output file as you go.
		int fileProbe = 0;
		while(fileProbe<files.size()){
		//Read in
		HSSFRow dataRow = output.createRow(rowCount);
		HSSFWorkbook readIn = new HSSFWorkbook(new FileInputStream(files.get(fileProbe)));
		HSSFWorkbook codedDataKey = new HSSFWorkbook(new FileInputStream("Coded Data.xls"));
		
		HSSFSheet readSheet = readIn.getSheetAt(0);
		HSSFRow readRow = readSheet.getRow(1);
		Cell instructor = readRow.getCell(3);
		Cell observer = readRow.getCell(2);
		DataFormatter formatter = new DataFormatter();
		
		HSSFCell ClassNumber = dataRow.createCell(6);
		ClassNumber.setCellValue(readRow.getCell(5).toString());
		
		//Get Class Size from Lous list somehow
		//HSSFCell ClassSize = dataRow.createCell(7);
		
		HSSFCell ClassLevel = dataRow.createCell(8);
		if((readRow.getCell(5).toString().substring(0,1).equals("1"))||(readRow.getCell(5).toString().substring(0,1).equals("2"))){
			ClassLevel.setCellValue("1");
		}
		if((readRow.getCell(5).toString().substring(0,1).equals("3"))||(readRow.getCell(5).toString().substring(0,1).equals("4"))){
			ClassLevel.setCellValue("2");
		}
		else{
			ClassLevel.setCellValue("3");
		}
		
		HSSFCell ClassDate = dataRow.createCell(9);
		ClassDate.setCellValue(readRow.getCell(6).toString());
		
		HSSFCell TotalTime = dataRow.createCell(10);
		int numRowsInput = readSheet.getPhysicalNumberOfRows();
		Date startTime = readSheet.getRow(1).getCell(1).getDateCellValue();
		Date endTime = readSheet.getRow(numRowsInput-1).getCell(1).getDateCellValue();
		long diffInMins = ((endTime.getTime() - startTime.getTime())/1000)/60;
		TotalTime.setCellValue(diffInMins);
		
		//Is this always 1?
		HSSFCell Observation = dataRow.createCell(11);
		Observation.setCellValue(1);
		
		HSSFCell Observer = dataRow.createCell(12);
		Observer.setCellValue(readRow.getCell(2).toString());
		
		
		//Use coded sheet to get ObserverID information
		HSSFCell ObserverID = dataRow.createCell(13);
		HSSFSheet observerSheet = codedDataKey.getSheetAt(2);
		int numRowsObservers = observerSheet.getPhysicalNumberOfRows();
		int numColsObservers = 2;
		for(int i = 0; i<numRowsObservers; i++){
			Row row = observerSheet.getRow(i);
			for (int j = 0; j < numColsObservers; j++) {
		        Cell cell = row.getCell(j);
		        String value = formatter.formatCellValue(cell);
		        if(value.equals(observer.toString())){
		        	ObserverID.setCellValue(row.getCell(1).toString());
		        }
			}
		}
		
		
		
		//Read the Code Sheet Data and match instructor information (Name, ID, Year, etc.) and append it to the output sheet.
		HSSFSheet codeSheet = codedDataKey.getSheetAt(0);
		int numRows = codeSheet.getPhysicalNumberOfRows();
		int numCols = codeSheet.getRow(0).getPhysicalNumberOfCells();
		for (int i = 0; i < numRows; i++) {
		    Row row = codeSheet.getRow(i);
		    for (int j = 0; j < numCols; j++) {
		        Cell cell = row.getCell(j);
		    	String value = formatter.formatCellValue(cell);
		      if(value.equals(instructor.toString())){
		        	HSSFCell instructorID = dataRow.createCell(0);
		        	instructorID.setCellValue(row.getCell(j+2).toString());
		        	HSSFCell instructorName = dataRow.createCell(1);
		        	instructorName.setCellValue(row.getCell(j).toString());
		        	HSSFCell classSubject = dataRow.createCell(2);
		        	classSubject.setCellValue(row.getCell(j+1).toString());
		        	HSSFCell treatment = dataRow.createCell(3);
		        	treatment.setCellValue(row.getCell(j+3).toString());
		        	HSSFCell program = dataRow.createCell(4);
		        	String programValue = formatter.formatCellValue(row.getCell(j+4));
		        	program.setCellValue(programValue);
		        	HSSFCell programYear = dataRow.createCell(5);
		        	String programYearValue = formatter.formatCellValue(row.getCell(j+5));
		        	programYear.setCellValue(programYearValue);
		        	break;
		       }
		    }
		}
		
		
		
		//Average every column of values from I-101 onward (ignoring the notes section) and append it to the output sheet
		int columnIndex = 10;
		int destLocation = 14;
		double sum = 0.0;
		double average;
		int numRowElements = readSheet.getPhysicalNumberOfRows()-1;
		for (int colIndex = columnIndex; colIndex <37; colIndex++){
			for (int rowIndex = 1; rowIndex<numRowElements+1; rowIndex++){
				if(colIndex!=23){ //I do not know how to handle the Notes column for averaging. So I skipped them.
				Row row = readSheet.getRow(rowIndex);
				Cell cell = row.getCell(colIndex);
				String Svalue = formatter.formatCellValue(cell);
				int usableValue = Integer.parseInt(Svalue);
				sum+=usableValue;
				}
			}
			average = (sum/numRowElements)*100;
			HSSFCell dest = dataRow.createCell(destLocation);
			dest.setCellValue(average);
			sum = 0.0;
			average = 0.0;
			destLocation++;
		}
		
		
		
		//iterate through the list of files and add a row to the output sheet
		fileProbe++;
		rowCount++;
		
		//close reading workbooks
		readIn.close();
		codedDataKey.close();
		}
		
		
		//Write and create the output file
		writenew.write(new FileOutputStream("output1.xls"));
		writenew.close();
		
		  }		
	}
	
	



