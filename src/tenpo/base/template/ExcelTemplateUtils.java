package tenpo.base.template;

import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author kyou
 * @link http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
 */
public class ExcelTemplateUtils {


	 /** DEFAULT CONSTRUCTOR. */
	 private ExcelTemplateUtils() {}

	 /**
	  * @param newSheet 目標シート
	  * @param sheet　コピー元シート
	  * @param copyStyle falseの場合、データのみをコピーする
	  */
	 public static void copySheets(final HSSFSheet newSheet, final HSSFSheet oldSheet, final boolean copyStyle){
	     int maxColumnNum = 0;
	     for (int i = oldSheet.getFirstRowNum(); i <= oldSheet.getLastRowNum(); i++) {
	         HSSFRow oldRow = oldSheet.getRow(i);
	         HSSFRow newRow = newSheet.createRow(i);
	         if (oldRow != null) {
	        	 copyRow(newRow, oldRow, copyStyle);
	             if (oldRow.getLastCellNum() > maxColumnNum) {
	                 maxColumnNum = oldRow.getLastCellNum();
	             }
	         }
	     }
	     for (int i = 0; i <= maxColumnNum; i++) {
	         newSheet.setColumnWidth(i, oldSheet.getColumnWidth(i));
	     }

	     if (copyStyle) {
	    	 copySheetSettings(newSheet, oldSheet);
	     }
	 }

	 /**
	  * @param oldRow コピー元シート
	  * @param newRow 目標シート
	  * @param copyStyle -
	  */
	 public static void copyRow(final HSSFRow newRow, final HSSFRow oldRow, final boolean copyStyle) {
		 HSSFSheet srcSheet = oldRow.getSheet();
		 HSSFSheet destSheet = newRow.getSheet();

		 int j = oldRow.getFirstCellNum();
		 if ( j<0 ) {j=0;}
		 int count = oldRow.getLastCellNum();

		 // manage a list of merged zone in order to not insert two times a merged zone
		 Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
	     newRow.setHeight(oldRow.getHeight());
	     // pour chaque row
	     for (; j <= count; j++) {
	         HSSFCell oldCell = oldRow.getCell(j);   // ancienne cell
	         HSSFCell newCell = newRow.getCell(j);  // new cell
	         if (oldCell != null) {
	             if (newCell == null) {
	                 newCell = newRow.createCell(j);
	             }
	             // copy chaque cell
	             copyCell(newCell, oldCell, copyStyle);
	             // copy les informations de fusion entre les cellules
	             //System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short)oldCell.getColumnIndex());
	             CellRangeAddress mergedRegion = getMergedRegion(srcSheet, oldRow.getRowNum(), (short)oldCell.getColumnIndex());

	             if (mergedRegion != null) {
	               //System.out.println("Selected merged region: " + mergedRegion.toString());
	               CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(), mergedRegion.getLastRow(), mergedRegion.getFirstColumn(),  mergedRegion.getLastColumn());
	                 //System.out.println("New merged region: " + newMergedRegion.toString());
	                 CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
	                 if (isNewMergedRegion(wrapper, mergedRegions)) {
	                     mergedRegions.add(wrapper);
	                     destSheet.addMergedRegion(wrapper.range);
	                 }
	             }
	         }
	     }
	 }

	 /**
	  * @param newCell 目標シート
	  * @param oldCell コピー元シート
	  * @param copyStyle -
	  */
	 public static void copyCell(final HSSFCell newCell, final HSSFCell oldCell, final boolean copyStyle) {
		 if(copyStyle) {
			 copyCellStyle(newCell, oldCell);
		 }
		 switch(oldCell.getCellType()) {
			 case HSSFCell.CELL_TYPE_STRING:
				 newCell.setCellValue(oldCell.getStringCellValue());
				 break;
			 case HSSFCell.CELL_TYPE_NUMERIC:
				 newCell.setCellValue(oldCell.getNumericCellValue());
				 break;
			 case HSSFCell.CELL_TYPE_BLANK:
				 newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
				 break;
			 case HSSFCell.CELL_TYPE_BOOLEAN:
				 newCell.setCellValue(oldCell.getBooleanCellValue());
				 break;
			 case HSSFCell.CELL_TYPE_ERROR:
				 newCell.setCellErrorValue(oldCell.getErrorCellValue());
				 break;
			 case HSSFCell.CELL_TYPE_FORMULA:
				 newCell.setCellFormula(oldCell.getCellFormula());
				 break;
			 default:
				 break;
		 }

	 }

	 /**
	  * @param oldRow コピー元シート
	  * @param newRow 目標シート
	  */
	 public static void copyRowStyle(final HSSFRow newRow, final HSSFRow oldRow) {
		 // manage a list of merged zone in order to not insert two times a merged zone
		 newRow.setHeight(oldRow.getHeight());
		 // pour chaque row
		 for (int j = oldRow.getFirstCellNum(); j <= oldRow.getLastCellNum(); j++) {
			 HSSFCell oldCell = oldRow.getCell(j);   // ancienne cell
			 HSSFCell newCell = newRow.getCell(j);  // new cell
			 if (oldCell != null) {
				 if (newCell == null) {
					 newCell = newRow.createCell(j);
				 }
				 // copy chaque cell
				 copyCellStyle(newCell, oldCell);
			 }
		 }

	 }

	 /**
	  * @param newCell 目標シート
	  * @param oldCell コピー元シート
	  */
	 public static void copyCellStyle(final HSSFCell newCell, final HSSFCell oldCell) {
         if(oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()){
             newCell.setCellStyle(oldCell.getCellStyle());
         } else{
             HSSFCellStyle newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
             newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
             newCell.setCellStyle(newCellStyle);
         }
	 }

	 /**
	  * @param newSheet 目標シート
	  * @param sheet　コピー元シート
	  */
	private static void copySheetSettings(final Sheet newSheet, final Sheet oldSheet) {

		newSheet.setAutobreaks(oldSheet.getAutobreaks());
		newSheet.setDefaultColumnWidth(oldSheet.getDefaultColumnWidth());
		newSheet.setDefaultRowHeight(oldSheet.getDefaultRowHeight());
		newSheet.setDefaultRowHeightInPoints(oldSheet.getDefaultRowHeightInPoints());
		newSheet.setDisplayGuts(oldSheet.getDisplayGuts());
		newSheet.setFitToPage(oldSheet.getFitToPage());

		newSheet.setForceFormulaRecalculation(oldSheet.getForceFormulaRecalculation());

		PrintSetup sheetToCopyPrintSetup = oldSheet.getPrintSetup();
		PrintSetup newSheetPrintSetup = newSheet.getPrintSetup();

		newSheetPrintSetup.setPaperSize(sheetToCopyPrintSetup.getPaperSize());
		newSheetPrintSetup.setScale(sheetToCopyPrintSetup.getScale());
		newSheetPrintSetup.setPageStart(sheetToCopyPrintSetup.getPageStart());
		newSheetPrintSetup.setFitWidth(sheetToCopyPrintSetup.getFitWidth());
		newSheetPrintSetup.setFitHeight(sheetToCopyPrintSetup.getFitHeight());
		newSheetPrintSetup.setLeftToRight(sheetToCopyPrintSetup.getLeftToRight());
		newSheetPrintSetup.setLandscape(sheetToCopyPrintSetup.getLandscape());
		newSheetPrintSetup.setValidSettings(sheetToCopyPrintSetup.getValidSettings());
		newSheetPrintSetup.setNoColor(sheetToCopyPrintSetup.getNoColor());
		newSheetPrintSetup.setDraft(sheetToCopyPrintSetup.getDraft());
		newSheetPrintSetup.setNotes(sheetToCopyPrintSetup.getNotes());
		newSheetPrintSetup.setNoOrientation(sheetToCopyPrintSetup.getNoOrientation());
		newSheetPrintSetup.setUsePage(sheetToCopyPrintSetup.getUsePage());
		newSheetPrintSetup.setHResolution(sheetToCopyPrintSetup.getHResolution());
		newSheetPrintSetup.setVResolution(sheetToCopyPrintSetup.getVResolution());
		newSheetPrintSetup.setHeaderMargin(sheetToCopyPrintSetup.getHeaderMargin());
		newSheetPrintSetup.setFooterMargin(sheetToCopyPrintSetup.getFooterMargin());
		newSheetPrintSetup.setCopies(sheetToCopyPrintSetup.getCopies());

		Header sheetToCopyHeader = oldSheet.getHeader();
		Header newSheetHeader = newSheet.getHeader();
		newSheetHeader.setCenter(sheetToCopyHeader.getCenter());
		newSheetHeader.setLeft(sheetToCopyHeader.getLeft());
		newSheetHeader.setRight(sheetToCopyHeader.getRight());

		Footer sheetToCopyFooter = oldSheet.getFooter();
		Footer newSheetFooter = newSheet.getFooter();
		newSheetFooter.setCenter(sheetToCopyFooter.getCenter());
		newSheetFooter.setLeft(sheetToCopyFooter.getLeft());
		newSheetFooter.setRight(sheetToCopyFooter.getRight());

		newSheet.setHorizontallyCenter(oldSheet.getHorizontallyCenter());
		newSheet.setMargin(Sheet.LeftMargin, oldSheet.getMargin(Sheet.LeftMargin));
		newSheet.setMargin(Sheet.RightMargin, oldSheet.getMargin(Sheet.RightMargin));
		newSheet.setMargin(Sheet.TopMargin, oldSheet.getMargin(Sheet.TopMargin));
		newSheet.setMargin(Sheet.BottomMargin, oldSheet.getMargin(Sheet.BottomMargin));

		newSheet.setPrintGridlines(oldSheet.isPrintGridlines());
		newSheet.setRowSumsBelow(oldSheet.getRowSumsBelow());
		newSheet.setRowSumsRight(oldSheet.getRowSumsRight());
		newSheet.setVerticallyCenter(oldSheet.getVerticallyCenter());
		newSheet.setDisplayFormulas(oldSheet.isDisplayFormulas());
		newSheet.setDisplayGridlines(oldSheet.isDisplayGridlines());
		newSheet.setDisplayRowColHeadings(oldSheet.isDisplayRowColHeadings());
		newSheet.setDisplayZeros(oldSheet.isDisplayZeros());
		newSheet.setPrintGridlines(oldSheet.isPrintGridlines());
		newSheet.setRightToLeft(oldSheet.isRightToLeft());
		newSheet.setZoom(1, 1);
     }

	 /**
	  * @link http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
	  * @param sheet the sheet containing the data.
	  * @param rowNum the num of the row to copy.
	  * @param cellNum the num of the cell to copy.
	  * @return the CellRangeAddress created.
	  */
	 private static CellRangeAddress getMergedRegion(final HSSFSheet sheet, final int rowNum, final short cellNum) {
	     for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
	         CellRangeAddress merged = sheet.getMergedRegion(i);
	         if (merged.isInRange(rowNum, cellNum)) {
	             return merged;
	         }
	     }
	     return null;
	 }

	 /**
	  * @link http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
	  * @param newMergedRegion the merged region to copy or not in the destination sheet.
	  * @param mergedRegions the list containing all the merged region.
	  * @return true if the merged region is already in the list or not.
	  */
	 private static boolean isNewMergedRegion(final CellRangeAddressWrapper newMergedRegion, final Set<CellRangeAddressWrapper> mergedRegions) {
	   return !mergedRegions.contains(newMergedRegion);
	 }
}


/**
 * @link http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
 */
class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {

    public CellRangeAddress range;

    /**
     * @param theRange the CellRangeAddress object to wrap.
     */
    public CellRangeAddressWrapper(final CellRangeAddress theRange) {
    	this.range = theRange;
    }

    /**
     * @param o the object to compare.
     * @return -1 the current instance is prior to the object in parameter, 0: equal, 1: after...
     */
    @Override
	public int compareTo(final CellRangeAddressWrapper o) {

        if (range.getFirstColumn() < o.range.getFirstColumn()
                    && range.getFirstRow() < o.range.getFirstRow()) {
              return -1;
        } else if (range.getFirstColumn() == o.range.getFirstColumn()
                    && range.getFirstRow() == o.range.getFirstRow()) {
              return 0;
        } else {
              return 1;
        }

    }

}
