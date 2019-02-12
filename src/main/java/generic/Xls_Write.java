package generic;




	import java.io.File;
	import java.io.FileOutputStream;
	import java.util.List;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.testng.Assert;

	public class Xls_Write {
		
		public static void createXLSFile(String excelFileName, String sheetName, int requiredRowCount, int requiredColumnCount,List<List<String>> data){
	    	System.out.println("Creating XLS file");
			try{
	    		XSSFWorkbook wb = new XSSFWorkbook();
	    		XSSFSheet sheet = wb.createSheet(sheetName);
	    		for (int i=0;i<requiredRowCount;i++){
	    			XSSFRow row = sheet.createRow(i);
	    			for(int j=0;j<requiredColumnCount;j++){
	    				XSSFCell cell = row.createCell(j);
	    				cell.setCellValue(data.get(i).get(j));
	    			}
	    		}
	    		FileOutputStream fileOut = new FileOutputStream(excelFileName);
	    		wb.write(fileOut);
	    		fileOut.flush();
	    		fileOut.close();
	    		File f=new File(excelFileName);
				Assert.assertTrue(f.exists() && !f.isDirectory());
				System.out.println("Creating XLS file succeed");
	    	}catch(Exception e){
	    		System.out.println("Creating XLS file failed");
	    		e.printStackTrace();
	    	}
	    }

		
		public static void createBlankXLSFile(String excelFileName, String sheetName){
		     System.out.println("Creating XLS file");
			try{
		     XSSFWorkbook wb = new XSSFWorkbook();
		     XSSFSheet sheet = wb.createSheet(sheetName);
		     FileOutputStream fileOut = new FileOutputStream(excelFileName);
		     wb.write(fileOut);
		     fileOut.flush();
		     fileOut.close();
		     File f=new File(excelFileName);
			Assert.assertTrue(f.exists() && !f.isDirectory());
			System.out.println("Creating XLS file succeed");
		     }catch(Exception e){
		     System.out.println("Creating XLS file failed");
		     e.printStackTrace();
		     }
		    }
		
		public static void main(String[] args) {
			// TODO Auto-generated method stub
			//createXLSFile("C:\\Users\\pranjan\\Desktop\\ABC.xlsx", "TESTINGXLS", 4, 5,"");

		}

	}



