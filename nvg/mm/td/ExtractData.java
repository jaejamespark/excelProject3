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
import nvg.mm.td.ModelNameRow;


public class ExtractData {
	@SuppressWarnings("null")
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

			// Put models into linkedlist when encountering different model name
			int numRows = sheet.getRows();
			
			String previousModelName =  null;
			LinkedList<ModelNameRow> modelList = new LinkedList<ModelNameRow>();
					
			for (int rows = 1; rows < numRows; ++rows) {
				Cell cellRows = sheet.getCell(colModel, rows);
				String modelName = cellRows.getContents();
				if (previousModelName != modelName)	{
					ModelNameRow modelNameRow = new ModelNameRow();
					modelNameRow.setModelName(modelName);
					modelNameRow.setModelRow(rows);
					modelList.add(modelNameRow);
					}
				previousModelName = modelName;
			}
			
			/* block out here */
			System.out.println("Total detected # of model name are " + modelList.size());
			for (int x = 0; x < modelList.size(); ++x){
				System.out.println("Current iteration is at " + x);
				System.out.println("Model Name is " + modelList.get(x).getModelName());
				System.out.println("Model Row is " + modelList.get(x).getModelRow());
			}
			/* block out here */
			
			
			// Go through linkedlist and remove duplicate
			LinkedList<Integer> modelDividerRow = new LinkedList<Integer>();
			modelDividerRow.add(1); // adding model in the 0 row
			for (int modelBaseCount = 0; modelBaseCount < modelList.size(); ++modelBaseCount){
				for (int nextModelCount = modelBaseCount + 1; nextModelCount < modelList.size(); ++nextModelCount){
					String nextModelName = modelList.get(nextModelCount).getModelName();
					String baseModelName = modelList.get(modelBaseCount).getModelName();
					if (!(baseModelName.regionMatches(0, nextModelName, 0, 5))){
						modelDividerRow.add(modelList.get(nextModelCount).getModelRow());
						modelBaseCount = nextModelCount - 1;
						break;
					}
				}
			}
			
			
			//System.out.println(modelDividerRow.get(0));
			
			
			
			System.out.println("end");
			
			//String testString = modelList.get(1).toString();
			//String subTestString = testString.substring(2, 4);
			
			//System.out.println(subTestString);
			
			
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
}
