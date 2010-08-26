package com.scut.zhili;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.InputStream;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class ReadXlsSheetsAndMerge {

	private Workbook output;
	private Sheet outputSheet;
	private int lastRowOfOutputSheet;
	public ReadXlsSheetsAndMerge() {
		try {
			InputStream inp = new FileInputStream("ZSRNC_utrancell_Model.xls");
			this.output = WorkbookFactory.create(inp);
			outputSheet = output.getSheetAt(0);
			lastRowOfOutputSheet = outputSheet.getLastRowNum() + 1;
		} catch (Exception e) {
			 System.err.println("where is the model file?");
		}
	}
	public static void setBorderStyle(CellStyle cellStyle){   
       
		  cellStyle.setBorderBottom(CellStyle.BORDER_THIN);   
	        cellStyle.setBorderTop(CellStyle.BORDER_THIN);   
	        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);   
	        cellStyle.setBorderRight(CellStyle.BORDER_THIN);   
	        //设置一个单元格边框颜色   
	       cellStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());   
	       cellStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());   
	       cellStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());   
	       cellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());           
    }   
	public static void CopyRightPrintOut() {
		System.out.println("        __    _ ___    __           ");
		System.out.println(" ____  / /_  (_) (_)  / /_  __  __  ");
		System.out.println("/_  / / __ \\/ / / /  / __ \\/ / / /  ");
		System.out.println(" / /_/ / / / / / /  / / / / /_/ /   ");
		System.out.println("/___/_/ /_/_/_/_/  /_/ /_/\\__,_/    ");
	}
	
	public boolean append(String file) {
		try {
			InputStream inp = new FileInputStream(file);
			Workbook other = WorkbookFactory.create(inp);
			CellStyle cellStyle = output.createCellStyle();
			cellStyle.setDataFormat(
					other.getCreationHelper().createDataFormat().getFormat("m/d/yy"));
			setBorderStyle(cellStyle);
			
		        
			CellStyle NumbercellStyle = output.createCellStyle();
			NumbercellStyle.setDataFormat(other.getCreationHelper().createDataFormat().getFormat("0.00"));
			setBorderStyle(NumbercellStyle);
			
			CellStyle basiccellStyle = output.createCellStyle();
			setBorderStyle(basiccellStyle);
			
			Sheet tmpSheet = other.getSheetAt(0); 
			//int lastRowOfOutputSheet = outputSheet.getLastRowNum() + 1;
			int lastRowOfOther = tmpSheet.getLastRowNum();
			//System.out.println(lastRowOfOther);
			if (tmpSheet.getPhysicalNumberOfRows() > 0) {
				Row Daterow = tmpSheet.getRow(4);
				Cell dateCellOfRecord = Daterow.getCell(3);
				
				Row RNCrow = tmpSheet.getRow(10);
				Cell rncCellOfRecord = RNCrow.getCell(3);
				
				if(DateUtil.isCellDateFormatted(dateCellOfRecord)) {
					//System.out.println(cell.getDateCellValue());

					for (int r = 14; r < lastRowOfOther-5; ++r) {
						int index = 0;
						Row rowTmp = tmpSheet.getRow(r);
						if (rowTmp == null) break;
						Row newRow = outputSheet.createRow(lastRowOfOutputSheet++);
						Cell firstColCellOfRow = newRow.createCell(index++);
						firstColCellOfRow.setCellStyle(cellStyle);
						firstColCellOfRow.setCellValue(dateCellOfRecord.getDateCellValue());
						Cell secondColCellOfRow = newRow.createCell(index++);
						secondColCellOfRow.setCellStyle(basiccellStyle);
						secondColCellOfRow.setCellValue(rncCellOfRecord.getStringCellValue());
						
						for (int col = 1; col < 79; ++col) {
							if (col == 2 || col == 3 || col == 5)
								continue;

							Cell cellTmp = rowTmp.getCell(col);
							if (cellTmp != null) {
								Cell newCell = newRow.createCell(index++);
								switch(cellTmp.getCellType()) {
							    	case Cell.CELL_TYPE_STRING:
							    		newCell.setCellStyle(basiccellStyle);
							    		newCell.setCellValue(cellTmp.getStringCellValue());
							    		break;
							    	case Cell.CELL_TYPE_NUMERIC:
							    		 //if(DateUtil.isCellDateFormatted(cellTmp)) {
							    			 //System.out.println(cellTmp.getDateCellValue());
							    	     //} else {
							    		newCell.setCellStyle(NumbercellStyle);
							    	    newCell.setCellValue(cellTmp.getNumericCellValue());
							    	    break;
							    	case Cell.CELL_TYPE_BOOLEAN:
							    		newCell.setCellValue(cellTmp.getBooleanCellValue());
							    		break;
							      	case Cell.CELL_TYPE_FORMULA:
							      		newCell.setCellValue(cellTmp.getCellFormula());
							      		break;
							      	case Cell.CELL_TYPE_BLANK:
							      		newCell.setCellValue(0.00);
							      		break;
							      	default:
							          System.out.println(cellTmp.getCellType());
								}
							}
						}
					}
				}
			}
			return true;
		} catch (Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows:.....");
            ex.printStackTrace(System.out);;
		}
		return false;
	}
	
	public void writeOutNewFile(String out) {	
	    // Write the output to a file 
		try {
			FileOutputStream fileOut = new FileOutputStream(out);
			output.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			System.out.println("output file failure!");
			e.printStackTrace();
		}
	}
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		CopyRightPrintOut();
		File srcFolder = new File("D:\\new");
		File[] filesList = null;
		if(srcFolder.isDirectory()) {
	            // Get a list of all of the Excel spreadsheet files (workbooks) in
	            // the source folder/directory
	            filesList = srcFolder.listFiles(new ExcelFilenameFilter());
	    } else {
	            // Assume that it must be a file handle - although there are other
	            // options the code should perhaps check - and store the reference
	            // into the filesList variable.
	            filesList = new File[]{srcFolder};
	    }
//		for (File f : filesList) {
//			System.out.println(f.getPath() + f.getName()); 
//		}
		ReadXlsSheetsAndMerge boReport = new ReadXlsSheetsAndMerge();
//		boReport.append("D:\\ZSRNC03_utrancell_0514.xls");
//		boReport.append("D:\\ZSRNC03_utrancell_0515.xls");
		for (File f : filesList) {
			boReport.append("D:\\new\\" + f.getName()); 
		}
		boReport.writeOutNewFile("UntranCellPerfBOReport.xls");
	}

}

class ExcelFilenameFilter implements FilenameFilter {

    public boolean accept(File file, String name) {
        return(name.endsWith(".xls"));// || name.endsWith(".xlsx"));
    }
}