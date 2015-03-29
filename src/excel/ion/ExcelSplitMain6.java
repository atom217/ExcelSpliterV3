package excel.ion;
/**
 * 
 */


import java.io.IOException;


/**
 * @author pushpendra.paliwal
 *
 */
public class ExcelSplitMain6 {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException {
			    ExcelOperations6 excelMaster2 = new ExcelOperations6();
			    if(args.length != 1) {
			    	 System.err.println("Invalid command line, exactly one argument required. If your file name is having spaces give the argument with double quotes.");
			    	  System.exit(1);
			    	}
			    String filename = args[0];
			    String extension = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
			    extension = extension.toUpperCase();
			   String excel = "XLSX";
			    if ((extension.equals(excel.toUpperCase())) == false) {
			    	System.err.println("Not valid file. Please give a valid excel file with extension xlsx");
			    	  System.exit(1);
			    }
			    excelMaster2.setInputFile(args[0]);
			    excelMaster2.readFile(filename);
		}
}

