package generic;


	
	
	
	


	import java.io.FileInputStream;
	import java.io.FileOutputStream;

	import java.util.ArrayList;

	import java.util.Iterator;
	import java.util.List;

	import org.apache.poi.common.usermodel.HyperlinkType;

	import org.apache.poi.hssf.usermodel.HSSFDateUtil;

	import org.apache.poi.ss.format.CellDateFormatter;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.CellStyle;

	import org.apache.poi.ss.usermodel.FillPatternType;

	import org.apache.poi.ss.usermodel.IndexedColors;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFCellStyle;
	import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
	import org.apache.poi.xssf.usermodel.XSSFDataValidation;
	import org.apache.poi.xssf.usermodel.XSSFFont;
	import org.apache.poi.xssf.usermodel.XSSFHyperlink;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;


	public class Xls_Reader
	{

	    public String path;
	    public FileInputStream fis = null;
	    public FileOutputStream fileOut = null;
	    private XSSFWorkbook workbook = null;
	    private XSSFSheet sheet = null;
	    private XSSFRow row = null;
	    private XSSFCell cell = null;
             

	
	    public Xls_Reader(String path)
	    {

		this.path = path;
		try
		{
		    // InputStream is = getClass().getResourceAsStream(path);
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);
		    sheet = workbook.getSheetAt(0);
		    fis.close();
		} catch (Exception e)
		{
		    // TODO Auto-generated catch block
		    e.printStackTrace();
		  

		}
	    }

	    
	    
	   
		   
	    // returns the row count in a sheet
	    public int getRowCount(String sheetName)
	    {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
		    return 0;
		else
		{
		    sheet = workbook.getSheetAt(index);
		    int number = sheet.getLastRowNum() + 1;
		    // System.out.println("Method : getRowCount:"+sheetName+" "+number);
		    return number;
		}

	    }

	 // returns the data from a cell
	    public String getCellData(String sheetName, String colName, int rowNum, int HeaderRow)
	    {

		try
		{
		
		    if (rowNum <= 0)
			return "";

		    int index = workbook.getSheetIndex(sheetName);
		    int col_Num = -1;
		    if (index == -1)
			return "";

		    sheet = workbook.getSheetAt(index);
		    if(HeaderRow>0)
		    	 row = sheet.getRow(HeaderRow);
		    else
		    	row = sheet.getRow(0);
		    
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim()
				.equals(colName.trim()))
			    col_Num = i;
		    }
		    if (col_Num == -1)
			return "";

		    sheet = workbook.getSheetAt(index);
		    row = sheet.getRow(rowNum - 1);
		    if (row == null)
			return "";
		    cell = row.getCell(col_Num);

		    if (cell == null)
			return "";

		    if (cell.getCellType() == Cell.CELL_TYPE_STRING)
			return cell.getStringCellValue();
		    else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
			    || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
		    {

			try{
				String cellText=null;
			if (HSSFDateUtil.isCellDateFormatted(cell))
			{
//		            below line is to evaluate the formulas.
			    workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
			  //  workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			    
			    
			    //date format as available in excel cell style data format
			    cellText = new CellDateFormatter(cell.getCellStyle().getDataFormatString()).format(cell.getDateCellValue()); 
		

		
			}else
			{  cell.setCellType(Cell.CELL_TYPE_STRING);
			    cellText = cell.getStringCellValue();
			}
			return cellText;
		    }
		    catch(Exception e1)
			{
			cell.setCellType(Cell.CELL_TYPE_STRING);
			return cell.getStringCellValue();
			
			}
			
			
		    } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
			return "";
		    else
			return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e)
		{

		    e.printStackTrace();
		    return "row " + rowNum + " or column " + colName
			    + " does not exist in xls";
		}
	    }
	    
	    // returns the data from a cell
	    public String getCellData(String sheetName, String colName, int rowNum)
	    {

		try
		{
		
		    if (rowNum <= 0)
			return "";

		    int index = workbook.getSheetIndex(sheetName);
		    int col_Num = -1;
		    if (index == -1)
			return "";

		    sheet = workbook.getSheetAt(index);
		    row = sheet.getRow(0);
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim()
				.equals(colName.trim()))
			    col_Num = i;
		    }
		    if (col_Num == -1)
			return "";

		    sheet = workbook.getSheetAt(index);
		    row = sheet.getRow(rowNum - 1);
		    if (row == null)
			return "";
		    cell = row.getCell(col_Num);

		    if (cell == null)
			return "";

		    if (cell.getCellType() == Cell.CELL_TYPE_STRING)
			return cell.getStringCellValue();
		    else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
			    || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
		    {

			try{
				String cellText=null;
			if (HSSFDateUtil.isCellDateFormatted(cell))
			{
//		            below line is to evaluate the formulas.
			    workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
			  //  workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			    
			    
			    //date format as available in excel cell style data format
			    cellText = new CellDateFormatter(cell.getCellStyle().getDataFormatString()).format(cell.getDateCellValue()); 
		
			
			}else
			{  cell.setCellType(Cell.CELL_TYPE_STRING);
			    cellText = cell.getStringCellValue();
			}
			return cellText;
		    }
		    catch(Exception e1)
			{
			cell.setCellType(Cell.CELL_TYPE_STRING);
			return cell.getStringCellValue();
			
			}
			
			
		    } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
			return "";
		    else
			return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e)
		{

		    e.printStackTrace();
		    return "row " + rowNum + " or column " + colName
			    + " does not exist in xls";
		}
	    }

	    // returns the data from a cell
	    
	    public String getCellData(String sheetName,int colNum,int rowNum){
	        try{
	              if(rowNum <=0){
	                    return "";
	              }
	              int index = workbook.getSheetIndex(sheetName);
	              if(index==-1){
	                    return "";
	              }
	              sheet = workbook.getSheetAt(index);
	              row = sheet.getRow(rowNum-1);
	              if(row==null){
	                    return "";
	              }
	              cell = row.getCell(colNum);
	              if(cell==null){
	                    return "";
	              }
	              if(cell.getCellType()==Cell.CELL_TYPE_STRING){
	                    return cell.getStringCellValue();
	              }
	              else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA){
	                    String cellText=null;
	                    try{
	                          if(HSSFDateUtil.isCellDateFormatted(cell)) {
	                                workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
	                                cellText = new CellDateFormatter(cell.getCellStyle().getDataFormatString()).format(cell.getDateCellValue());
	                        }else{
	                            cell.setCellType(Cell.CELL_TYPE_STRING);
	                            cellText = cell.getStringCellValue();
	                        }
	                          return cellText;
	                    }catch(Throwable t){
	                          cell.setCellType(Cell.CELL_TYPE_STRING);
	                    cellText = cell.getStringCellValue();
	                    return cellText;
	                    }
	                    
	              }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK){
	                      return "";
	              }
	            else{
	              return String.valueOf(cell.getBooleanCellValue());
	            }
	        }
	        catch(Exception e){
	              e.printStackTrace();
	              return "row "+rowNum+" or column "+colNum +" does not exist  in xls";
	        }
	  }

	    // returns true if data is set successfully else false

	    public boolean setCellDataRW(String sheetName, String colName, int rowNum,
		    String data)
	    {
		try
		{
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);

		    if (rowNum <= 0)
			return false;

		    int index = workbook.getSheetIndex(sheetName);
		    int colNum = -1;
		    if (index == -1)
			return false;

		    sheet = workbook.getSheetAt(index);

		    row = sheet.getRow(0);
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim().equals(colName))
			    colNum = i;
		    }
		    if (colNum == -1)
			return false;

		    sheet.autoSizeColumn(colNum);
		    row = sheet.getRow(rowNum - 1);
		    if (row == null)
			row = sheet.createRow(rowNum - 1);

		    cell = row.getCell(colNum);
		    if (cell == null)
			cell = row.createCell(colNum);

		    // cell style
		    // CellStyle cs = workbook.createCellStyle();
		    // cs.setWrapText(true);
		    // cell.setCellStyle(cs);
		    cell.setCellValue(data);
		    workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

		    fileOut = new FileOutputStream(path);

		    workbook.write(fileOut);

		    fileOut.close();

		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    public boolean setCellDataRW(String sheetName, int colNum, int rowNum,
	    	    String data)
	        {
	    	try
	    	{
	    	    fis = new FileInputStream(path);
	    	    workbook = new XSSFWorkbook(fis);

	    	    if (rowNum <= 0)
	    		return false;

	    	    int index = workbook.getSheetIndex(sheetName);
	    	    
	    	    if (index == -1)
	    		return false;

	    	    sheet = workbook.getSheetAt(index);

	    	    row = sheet.getRow(0);
//	    	    for (int i = 0; i < row.getLastCellNum(); i++)
//	    	    {
//	    		// System.out.println(row.getCell(i).getStringCellValue().trim());
//	    		if (row.getCell(i).getStringCellValue().trim().equals(colName))
//	    		    colNum = i;
//	    	    }
	    	    if (colNum == -1)
	    		return false;

	    	    sheet.autoSizeColumn(colNum);
	    	    row = sheet.getRow(rowNum - 1);
	    	    if (row == null)
	    		row = sheet.createRow(rowNum - 1);

	    	    cell = row.getCell(colNum);
	    	    if (cell == null)
	    		cell = row.createCell(colNum);

	    	    // cell style
	    	    // CellStyle cs = workbook.createCellStyle();
	    	    // cs.setWrapText(true);
	    	    // cell.setCellStyle(cs);
	    	    cell.setCellValue(data);
	    	    workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

	    	    fileOut = new FileOutputStream(path);

	    	    workbook.write(fileOut);

	    	    fileOut.close();

	    	} catch (Exception e)
	    	{
	    	    e.printStackTrace();
	    	    return false;
	    	}
	    	return true;
	        }
	    
	    public XSSFWorkbook openXls()
	    {
		try
		{
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);
		} catch (Exception e)
		{
		    e.printStackTrace();
		    // return false;
		}
		return workbook;
	    }

	    public boolean writeXls()
	    {

		try
		{
		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);
		    fileOut.close();
		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    // returns true if data is set successfully else false
	    public boolean setCellData(String sheetName, String colName, int rowNum,
		    String data)
	    {
		try
		{

		    if (rowNum <= 0)
			return false;

		    int index = workbook.getSheetIndex(sheetName);
		    int colNum = -1;
		    if (index == -1)
			return false;

		    sheet = workbook.getSheetAt(index);

		    row = sheet.getRow(0);
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim().equals(colName))
			    colNum = i;
		    }
		    if (colNum == -1)
			return false;

		    sheet.autoSizeColumn(colNum);
		    row = sheet.getRow(rowNum - 1);
		    if (row == null)
			row = sheet.createRow(rowNum - 1);

		    cell = row.getCell(colNum);
		    if (cell == null)
			cell = row.createCell(colNum);

		    // cell style
		    // CellStyle cs = workbook.createCellStyle();
		    // cs.setWrapText(true);
		    // cell.setCellStyle(cs);
		    cell.setCellValue(data);

		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    // returns true if data is set successfully else false

	    public boolean setCellData(String sheetName, String colName, int rowNum,
		    String data, String url)
	    {
		// System.out.println("setCellData setCellData******************");
		try
		{
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);

		    if (rowNum <= 0)
			return false;

		    int index = workbook.getSheetIndex(sheetName);
		    int colNum = -1;
		    if (index == -1)
			return false;

		    sheet = workbook.getSheetAt(index);
		    // System.out.println("A");
		    row = sheet.getRow(0);
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim()
				.equalsIgnoreCase(colName))
			    colNum = i;
		    }

		    if (colNum == -1)
			return false;
		    sheet.autoSizeColumn(colNum);
		    row = sheet.getRow(rowNum - 1);
		    if (row == null)
			row = sheet.createRow(rowNum - 1);

		    cell = row.getCell(colNum);
		    if (cell == null)
			cell = row.createCell(colNum);

		    cell.setCellValue(data);
		    XSSFCreationHelper createHelper = workbook.getCreationHelper();

		    // cell style for hyperlinks
		    // by default hypelrinks are blue and underlined
		    CellStyle hlink_style = workbook.createCellStyle();
		    XSSFFont hlink_font = workbook.createFont();
		    hlink_font.setUnderline(XSSFFont.U_SINGLE);
		    hlink_font.setColor(IndexedColors.BLUE.getIndex());
		    hlink_style.setFont(hlink_font);
		    // hlink_style.setWrapText(true);

		//    XSSFHyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
		 //   link.setAddress(url);
		   // cell.setHyperlink(link);
		    cell.setCellStyle(hlink_style);

		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);

		    fileOut.close();

		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    // returns true if sheet is created successfully else false
	    public boolean addSheet(String sheetname)
	    {

		FileOutputStream fileOut;
		try
		{
		    workbook.createSheet(sheetname);
		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);
		    fileOut.close();
		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    // returns true if sheet is removed successfully else false if sheet does
	    // not exist
	    public boolean removeSheet(String sheetName)
	    {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
		    return false;

		FileOutputStream fileOut;
		try
		{
		    workbook.removeSheetAt(index);
		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);
		    fileOut.close();
		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;
	    }

	    // returns true if column is created successfully
	    public boolean addColumn(String sheetName, String colName)
	    {
		// System.out.println("**************addColumn*********************");

		try
		{
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);
		    int index = workbook.getSheetIndex(sheetName);
		    if (index == -1)
			return false;

		    CellStyle style = workbook.createCellStyle();
		    style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		 //   style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		    sheet = workbook.getSheetAt(index);

		    row = sheet.getRow(0);
		    if (row == null)
			row = sheet.createRow(0);

		    // cell = row.getCell();
		    // if (cell == null)
		    // System.out.println(row.getLastCellNum());
		    if (row.getLastCellNum() == -1)
			cell = row.createCell(0);
		    else
			cell = row.createCell(row.getLastCellNum());

		    cell.setCellValue(colName);
		    cell.setCellStyle(style);

		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);
		    fileOut.close();

		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}

		return true;

	    }

	    // removes a column and all the contents
	    public boolean removeColumn(String sheetName, int colNum)
	    {
		try
		{
		    if (!isSheetExist(sheetName))
			return false;
		    fis = new FileInputStream(path);
		    workbook = new XSSFWorkbook(fis);
		    sheet = workbook.getSheet(sheetName);
		    XSSFCellStyle style = workbook.createCellStyle();
		    style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		    XSSFCreationHelper createHelper = workbook.getCreationHelper();
		    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		    for (int i = 0; i < getRowCount(sheetName); i++)
		    {
			row = sheet.getRow(i);
			if (row != null)
			{
			    cell = row.getCell(colNum);
			    if (cell != null)
			    {
				cell.setCellStyle(style);
				row.removeCell(cell);
			    }
			}
		    }
		    fileOut = new FileOutputStream(path);
		    workbook.write(fileOut);
		    fileOut.close();
		} catch (Exception e)
		{
		    e.printStackTrace();
		    return false;
		}
		return true;

	    }

	    // find whether sheets exists
	    public boolean isSheetExist(String sheetName)
	    {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
		{
		    index = workbook.getSheetIndex(sheetName.toUpperCase());
		    if (index == -1)
			return false;
		    else
			return true;
		} else
		    return true;
	    }

	    // returns number of columns in a sheet
	    public int getColumnCount(String sheetName)
	    {
		// check if sheet exists
		if (!isSheetExist(sheetName))
		    return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);

		if (row == null)
		    return -1;

		return row.getLastCellNum();

	    }

	    // String sheetName, String testCaseName,String keyword ,String URL,String
	    // message
	    public boolean addHyperLink(String sheetName, String screenShotColName,
		    String testCaseName, int index, String url, String message)
	    {
		// System.out.println("ADDING addHyperLink******************");

		url = url.replace('\\', '/');
		if (!isSheetExist(sheetName))
		    return false;

		sheet = workbook.getSheet(sheetName);

		for (int i = 2; i <= getRowCount(sheetName); i++)
		{
		    if (getCellData(sheetName, 0, i).equalsIgnoreCase(testCaseName))
		    {
			// System.out.println("**caught "+(i+index));
			setCellData(sheetName, screenShotColName, i + index, message,
				url);
			break;
		    }
		}

		return true;
	    }

	    public int getCellRowNum(String sheetName, String colName, String cellValue)
	    {

		for (int i = 2; i <= getRowCount(sheetName); i++)
		{
		    if (getCellData(sheetName, colName, i).equalsIgnoreCase(cellValue))
		    {
			return i;
		    }
		}
		return -1;

	    }

	    public int getCellRowNum(String sheetName, int colNum, String cellValue)
	    {

		for (int i = 2; i <= getRowCount(sheetName); i++)
		{
		    if (getCellData(sheetName, colNum, i).equalsIgnoreCase(cellValue))
		    {
			return i;
		    }
		}
		return -1;

	    }
	    public boolean shiftRow(String sheetName, int startFrom, int numberOfRow,
		    int extraCount)
	    { // extraCount if we want ot make some adjustment in the count

		try
		{
		    fileOut = new FileOutputStream(path);

		} catch (Exception e)
		{

		}

		try
		{
		    if (!isSheetExist(sheetName))
			return false;

		    sheet = workbook.getSheet(sheetName);

		    sheet.shiftRows(startFrom, sheet.getLastRowNum() + extraCount,
			    numberOfRow);
		    workbook.write(fileOut);
		} catch (Exception e)
		{

		}
		return true;
	    }

	    public int getLastRowNum(String sheetName)
	    {
		int rowCount;
		// check if sheet exists
		if (!isSheetExist(sheetName))
		    return -1;

		sheet = workbook.getSheet(sheetName);
		rowCount = sheet.getLastRowNum();
		
		//if (rowCount == 0) {
			rowCount = this.sheet.getPhysicalNumberOfRows();
		//}
		return rowCount;
	    }

	  
	    
	    // sheet.setAutoFilter(new CellRangeAddress(firstCell.getRow(),
	    // lastCell.getRow(), firstCell.getCol(), lastCell.getCol()));
	    // my_sheet.setAutoFilter(CellRangeAddress.valueOf("A1:C5"));

	    public boolean checkColExist(String resColName, String sheetName,
	    	    Xls_Reader lclCurrentSuiteXls)
	        {

	    	for (int i = 0; i <= lclCurrentSuiteXls.getColumnCount(sheetName); i++)
	    	{

	    	    String colHeading = lclCurrentSuiteXls.getCellData(sheetName, i, 1);

	    	    if (colHeading.equals(resColName))
	    	    {
	    		return true;
	    	    }
	    	}

	    	return false;

	        }
	    
	    public boolean isCellBlank(Cell c) {
	  		return (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK);
	  	}
	    
	      public boolean isCellEmpty(Cell c) {
	  		return (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK || (c
	  				.getCellType() == Cell.CELL_TYPE_STRING && c
	  				.getStringCellValue().isEmpty()));
	  	}
	      
	    public int getlastCellNum(String sheetName, String colName)
	    {
	 // Get the last cell number of a column
	    	
		try
		{
		
		 
		    int index = workbook.getSheetIndex(sheetName);
		    int col_Num = -1;
		    if (index == -1)
			return index;

		    sheet = workbook.getSheetAt(index);
		    row = sheet.getRow(0);
		    for (int i = 0; i < row.getLastCellNum(); i++)
		    {
			// System.out.println(row.getCell(i).getStringCellValue().trim());
			if (row.getCell(i).getStringCellValue().trim()
				.equals(colName.trim()))
			    col_Num = i;
		    }
		    if (col_Num == -1)
			return col_Num;

		    sheet = workbook.getSheetAt(index);
		    int Rowcount=sheet.getLastRowNum();
		    
		    for(int j=Rowcount; j>=0;j--)
		    {
		    row = sheet.getRow(j);
		    cell = row.getCell(col_Num);
		    if(!isCellBlank(cell))
			return j+1;
		    }
		    return -1;

		} catch (Exception e)
		{

		    e.printStackTrace();
		    return -1;
		}
		
	    }
	    
	    public int getlastCellNum(String sheetName, int colNum)
	    {
	 // Get the last cell number of a column
	    	
		try
		{
		    int index = workbook.getSheetIndex(sheetName);
		    if (index == -1)
			return index;

		    sheet = workbook.getSheetAt(index);
		    row = sheet.getRow(0);
		   
		    if (colNum == -1)
			return colNum;

		    sheet = workbook.getSheetAt(index);
		    int Rowcount=sheet.getLastRowNum();
		    
		    for(int j=Rowcount; j>=0;j--)
		    {
		    row = sheet.getRow(j);
		    cell = row.getCell(colNum);
		    if(!isCellBlank(cell))
			return j+1;
		    }
		    return -1;

		} catch (Exception e)
		{

		    e.printStackTrace();
		    return -1;
		}
		
	    }
	    public Boolean verifyDropDownList(String sheetname, String Columnname, String Columnvalue) {
	       	
	           String[] ColumnListValues=Columnvalue.split("\\`");
	        	int index = workbook.getSheetIndex(sheetname);
	            XSSFSheet sheet = workbook.getSheetAt(index);
	        	ArrayList<String> sheetListContents =new   ArrayList<String>();
	        	List<XSSFDataValidation> ls=sheet.getDataValidations();
	        	XSSFSheet ddListSheet = workbook.getSheet("column"+Columnname);
	        	
	        	Iterator<Row> iterator = ddListSheet.iterator();
	            
	            while (iterator.hasNext()) {
	                Row nextRow = iterator.next();
	                Iterator<Cell> cellIterator = nextRow.cellIterator();
	                 
	                while (cellIterator.hasNext()) {
	                    Cell cell = cellIterator.next();
	                    System.out.println(cell.getStringCellValue());
	                    sheetListContents.add(cell.getStringCellValue().toLowerCase());
	                }
	               
	            }
	        	
	            for(String key:ColumnListValues){
	            	if(!sheetListContents.contains(key.toLowerCase())){
	            		return false;
	            	}
	            }
	            
	             return true;
	        }
	    
	}


	
