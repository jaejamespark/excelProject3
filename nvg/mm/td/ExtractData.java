package nvg.mm.td;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.ListIterator;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.biff.RecordData;

public class ExtractData {
	public static void main(String[] args) {
		
		try {
			// Read a workbook from a file
			Workbook workbook = Workbook.getWorkbook(new File("d631_all.xls"));
			
			// Read a sheet from the workbook
			Sheet sheet = workbook.getSheet(0);
			
			// Identify number of columns in the sheet
			int numCols = sheet.getColumns();
			
			// Find columns of Model, Status, Priority, Detected in version
			int colModel = 0, colStatus = 0, colPriority = 0, colDetectVer = 0;
			
			for (int cols = 0; cols < numCols; ++cols) {
				Cell cellCols = sheet.getCell(cols, 0);
				
				String label = cellCols.getContents();
				
				if (label.equals("Model")) {
					colModel = cols;
					System.out.println("Model is at col " + colModel);
				}
				
				if (label.equals("Status")) {
					colStatus = cols;
					System.out.println("Status is at col " + colStatus);
				}
				
				if (label.equals("Priority")) {
					colPriority = cols;
					System.out.println("Priority is at col " + colPriority);
				}
				
				if (label.equals("Detected in Version")) {
					colDetectVer = cols;
					System.out.println("Detected in Version is at col " + colDetectVer);
				}
			}

			// Identify Models and number of Models
			int numRows = sheet.getRows();
			
			String previousModelName =  null;
			LinkedList modelList = new LinkedList();
			ListIterator modelListIterator = modelList.listIterator();
					
			for (int rows = 1; rows < numRows; ++rows) {
				Cell cellRows = sheet.getCell(colModel, rows);
				String modelName = cellRows.getContents();
				if (previousModelName != modelName)	{
					
					// If there is no match with already existing linkedlist
					while (rows == 1 || modelListIterator.hasNext()) {
						if (modelListIterator.next() == modelName) {
							break;
						}
						modelList.add(modelName);
					}
					
					
					}
				previousModelName = modelName;
			}
			System.out.println(modelList);
			
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
}
