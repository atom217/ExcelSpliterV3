package excel.ion;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations6 {

	  private String inputFile;
	  private String fileName;
	  public void setInputFile(String inputFile) {
	    this.inputFile = inputFile;
	  }
	  
	  public void readFile(String inputFilename)
	    {
	        try
	        {
		            FileInputStream file = new FileInputStream(new File(inputFile)); //input file
		            XSSFWorkbook workbook = new XSSFWorkbook(file); //input workbook
		            XSSFSheet sheet = workbook.getSheetAt(0);  //get 0th sheet
		            Iterator<Row> rowIterator = sheet.iterator(); // row iteration 
		            
		            List<CellStyle> cellStyleList = new ArrayList<CellStyle>();
		            List<CellStyle> headerStyleList = new ArrayList<CellStyle>();
		            List<String> header = new ArrayList<String>();
		            List<String> cellList = new ArrayList<String>();
		            while (rowIterator.hasNext())
		            {
		                //create new workbook define sheet and row and cells
		            	XSSFWorkbook writeToWorkBook = new XSSFWorkbook(); //output for each workbook
		            	CreationHelper createHelper = writeToWorkBook.getCreationHelper();
		            	XSSFSheet writeToSheet = writeToWorkBook.createSheet("Main");// Common name for all output sheets
		            	Row writeToRow = null;
		                Cell writeToCell = null;
		            	Row row = rowIterator.next(); //For each row, iterate through all the columns
		                
		            	if(row.getRowNum() != 0) // don't read headers . Reading Headers in else statement
		                {
		            		writeToRow = writeToSheet.createRow(1); //1st Row created HARD CODED
			                Iterator<Cell> cellIterator = row.cellIterator(); //read cells
			                int currentCellNumber = 0; 
			                String tempFileName1 = null;
			                String tempFileName3 = null;
			                //String tempFileName10 = null; //variable not in use //28 03 
			                while (cellIterator.hasNext())
			                {
			                    Cell cell = cellIterator.next();
			                    writeToCell = writeToRow.createCell(currentCellNumber); // new cell created
			                    
			                 switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_BLANK:
                                	writeToCell.setCellType(Cell.CELL_TYPE_BLANK);
                                	writeToCell.setCellValue("");
                                    break;

                                case Cell.CELL_TYPE_BOOLEAN:
                                	writeToCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                                	writeToCell.setCellValue(cell.getBooleanCellValue());
                                    break;

                                case Cell.CELL_TYPE_ERROR:
                                	writeToCell.setCellType(Cell.CELL_TYPE_ERROR);
                                	writeToCell.setCellErrorValue(cell.getErrorCellValue());
                                    break;

                                case Cell.CELL_TYPE_FORMULA:
                                	writeToCell.setCellType(Cell.CELL_TYPE_FORMULA);
                                	writeToCell.setCellFormula(cell.getCellFormula());
                                    break;

                                case Cell.CELL_TYPE_NUMERIC:
                                	if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                		 Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                                		 String strValue = new CellDateFormatter(cell.getCellStyle().getDataFormatString()).format(date);
                                		 writeToCell.setCellValue(strValue);
                                	}else{
                                		writeToCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                                		writeToCell.setCellValue(new Double(cell.getNumericCellValue()));
                                		
                                	}
                                    break;

                                case Cell.CELL_TYPE_STRING:
                                	writeToCell.setCellType(Cell.CELL_TYPE_STRING);
                                	writeToCell.setCellValue(cell.getStringCellValue());
                                	Hyperlink linkAddress = cell.getHyperlink();
                                	String inputFilename2 = inputFilename ;
                                	if(linkAddress != null){
                                		System.out.println(linkAddress.getAddress().toString());
                                		File f2 = new File(inputFilename2);
                    		    	 	System.out.println(f2.getName());
                    		    	 	inputFilename2 = inputFilename2.replace(f2.getName(),"");
                    		    	 	inputFilename2 = inputFilename2.replace("\\","/");
                    		    	 	String latestFileName = "file:///"+inputFilename2 + cell.getHyperlink().getAddress();
                    		    	 	//latestFileName = URLEncoder.encode(latestFileName, "UTF-8"); 
                                		linkAddress.setAddress(latestFileName);
                                		XSSFHyperlink updatedLinkAdderess = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_FILE);
										updatedLinkAdderess.setAddress(latestFileName);
                                		writeToCell.setHyperlink(updatedLinkAdderess);
                                	}
                                	XSSFCellStyle newStyle = writeToWorkBook.createCellStyle(); //create a style to clone
                                	newStyle.cloneStyleFrom(cell.getCellStyle()); //cloning
                                	newStyle.getCoreXf().unsetBorderId();  //to get rid of error while opening excel sheet 
                                	newStyle.getCoreXf().unsetFillId(); //to get rid of error while opening excel sheet 
                                	writeToCell.setCellStyle(newStyle);
                                    break;
                                default:
                                	writeToCell.setCellFormula(cell.getCellFormula());
                                } //switch statement
			                 /*
			                  tempFileName = cellList.get(3)+"_"+cellList.get(1);
			       		      System.out.print("/////"+ tempFileName + "//////");
			       		      setFileName(tempFileName);*/
			                 if (currentCellNumber == 1){
			                  cell.setCellType(Cell.CELL_TYPE_STRING);
			                  tempFileName1 = cell.getStringCellValue();
			                  tempFileName1 = tempFileName1.replaceAll("[^a-zA-Z0-9]", "_"); //removing special character
			                 }
			                 if(currentCellNumber == 3){
			                	 cell.setCellType(Cell.CELL_TYPE_STRING);
			                	 tempFileName3 = cell.getStringCellValue();
				                 tempFileName3 = tempFileName3.replaceAll("[^a-zA-Z0-9]", "_"); //removing special character
			                 }
			                 
			                 currentCellNumber++;
			               // System.out.println("\n");
		                }  //end of while cell iterations
			                if (tempFileName3 == null){
			                	tempFileName3 = "";
			                }
			                if (tempFileName1 == null){
			                	tempFileName1 = "";
			                }
			                String tempfileName = tempFileName3 +"_"+tempFileName1;//+"_"+ tempFileName10;
			                writeToNewWorkbook(writeToWorkBook,header,headerStyleList,inputFilename,tempfileName);
			                tempfileName = null;
			                cellList.clear();
			                cellStyleList.clear();
		                
		                } //  if(row.getRowNum() != 0)
		                else{
			                Iterator<Cell> cellIterator = row.cellIterator();
			                while (cellIterator.hasNext())
			                {
			                	Cell cell = cellIterator.next();
			                    //Check the cell type and format accordingly
			                    switch (cell.getCellType())
			                    {
			                    
			                    	case Cell.CELL_TYPE_BLANK:
			                    			headerStyleList.add(cell.getCellStyle());	
			                    			header.add("")	;
	                            		break;
			                    	case Cell.CELL_TYPE_NUMERIC:
			                            		headerStyleList.add(cell.getCellStyle());	
			                            		header.add("" + cell.getNumericCellValue())	;
			                            break;
			                        case Cell.CELL_TYPE_STRING:
			                            		headerStyleList.add(cell.getCellStyle());	
			                            		header.add(cell.getStringCellValue())	;
			         			         break;
			                        case Cell.CELL_TYPE_BOOLEAN:
	                            				headerStyleList.add(cell.getCellStyle());	
	                            				header.add(""+cell.getBooleanCellValue())	;
	                            		break;
			                       
			                    }//switch statement 
			                }//while statement 
			                System.out.println("\n");
		                } //end of if (reading headers)
		            } //end of while all row iterations
		            System.out.println("\n\n\n\n\n !!!!Congratulations File Successfully Splited !!!!!!!!!!!!!");
		            file.close();
		                
	        }catch (Exception e)
	        {
	            e.printStackTrace();
	        }
			
	    }

		private void writeToNewWorkbook(XSSFWorkbook writeToWorkBook,List<String> headList, List<CellStyle> headerStyleList,String inputFilename, String tempFileName) {
			XSSFSheet writeToSheet = writeToWorkBook.getSheetAt(0);
			System.out.print("/////"+ tempFileName + "//////");
		    setFileName(tempFileName);
		    
		    Map<String, Object[]> data = new TreeMap<String, Object[]>();
		      int y = headList.size();
		      String []headArray = new String[y];
		      headArray= headList.toArray(headArray);
		      data.put("1", headArray);
		      Set<String> keyset = data.keySet();
			  for (String key : keyset)
			  {
			        Row nrow = writeToSheet.createRow(0);
			        Object [] objArr = data.get(key);
			        int cellnum = 0;
			        for (Object obj : objArr)
			        {
			           Cell ncell = nrow.createCell(cellnum);
			           if(obj instanceof String)
			                {ncell.setCellValue((String)obj);}
			            else if(obj instanceof Integer)
			                {ncell.setCellValue((Integer)obj);}
			          
			           if(key.equals("1")){
			        	 XSSFCellStyle newHeaderStyle = writeToWorkBook.createCellStyle();
			        	 newHeaderStyle.cloneStyleFrom(headerStyleList.get(cellnum));
			        	 newHeaderStyle.getCoreXf().unsetBorderId();  //to get rid of error while opening excel sheet 
			        	 newHeaderStyle.getCoreXf().unsetFillId(); //to get rid of error while opening excel sheet 
		       	    	 ncell.setCellStyle(newHeaderStyle);
			           }
			           cellnum++;
			        }
			     }
			     
		    try{
		    	if(!(tempFileName.equals("_")))
		    	{
		    		System.out.println("\n\n\\n\n !!!!!!!!!@@@@@"+inputFilename);
		    	 	File f = new File(inputFilename);
		    	 	System.out.println(f.getName());
		    	 	inputFilename = inputFilename.replace(f.getName(),"");
		    	 	String latestFileName = inputFilename + getFileName()+".xlsx";
		    	 	File outputFile = new File(latestFileName);
		    	 	int fileCount = 1; 
		    	 	while(outputFile.exists()) {
		    	 		latestFileName = inputFilename + getFileName()+"_"+ Integer.toString(fileCount++)+".xlsx";
		    	 		outputFile = new File(latestFileName);
		    	 	}//if statement 
		    	 	FileOutputStream out = new FileOutputStream(outputFile);
		    		writeToWorkBook.write(out);
		    		System.out.println("\n\n\n\n\n !!!!!!!!!!!\n"+getFileName() + "written successfully on disk.");
		    		out.close();
		    	}//if	
		    }//try
		    catch (Exception e){
		        e.printStackTrace();
		    }//catch

		}
		public String getFileName() {
			return fileName;
		}
		public void setFileName(String fileName){
			this.fileName = fileName;
		}
		
}
