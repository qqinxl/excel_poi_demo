package tenpo.base.old;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.util.CellRangeAddress;

public class RptExcelWrapper {

	protected FileOutputStream output;
	protected HSSFWorkbook book;
	// 新規シートを追加用
	protected Map<String, HSSFSheet> sheetMap = new HashMap<String, HSSFSheet>();
	// 現時点操作しているシート
	protected HSSFSheet sheet;
	protected HSSFRow row;
	protected HSSFCell cell;

	protected int rowNo;
	protected int cellNo;

	public RptExcelWrapper(final String fileName, final String sheetName){
		try{
			book= new HSSFWorkbook();
			newSheet(sheetName);
			output = new FileOutputStream(fileName);

		}catch(Exception e){
			e.printStackTrace();
		}
	}
	// isLandscapeがtrueの場合、横Print
	public void setLandscape(final boolean isLandscape) {
		HSSFPrintSetup printSetup = sheet.getPrintSetup();
	    printSetup.setLandscape(isLandscape);
	    printSetup.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);
	}
	// 0:Left; 1:center; 2:right
	public void setHeader(final String msg, final HeaderFooterLocation location) {
		//Headerの設定
        Header header= sheet.getHeader();
        if (location.equals(HeaderFooterLocation.LEFT)) {
        	header.setLeft(msg);
        } else if (location.equals(HeaderFooterLocation.CENTER)) {
        	header.setCenter(msg);
        } else {
        	header.setRight(msg);
        }
	}
	public void setFooter(final String msg, final HeaderFooterLocation location) {
		//Headerの設定
		Footer footer = sheet.getFooter();
		if (location.equals(HeaderFooterLocation.LEFT)) {
			footer.setLeft(msg);
		} else if (location.equals(HeaderFooterLocation.CENTER)) {
			footer.setCenter(msg);
		} else {
			footer.setRight(msg);
		}
	}
    // 新規シートを追加する
	public void newSheet(final String sheetName) {
		if (sheetMap.containsKey(sheetName)) {
			sheet = sheetMap.get(sheetName);
		} else {
			sheet = book.createSheet(sheetName);
			sheetMap.put(sheetName, sheet);
		}
		row = null;
		cell = null;
		rowNo = 0;
		cellNo = 0;
	}
	// 行を合併する
	public void mergedRegion(final int y1, final int y2, final int x1, final int x2) {
		sheet.addMergedRegion(new CellRangeAddress(y1, y2, x1, x2));
	}
	public HSSFCellStyle createStyle(final String styleName){
		HSSFCellStyle style = book.createCellStyle();

		if(styleName.equals("aaa")){
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		}

		return style;
	}

	public void setColumnWidth(final int column,final int size){
		//TODO
	}

	public void newLine(){
		try{
			row = sheet.createRow(rowNo++);
			cellNo =0;
		}catch(Exception e){
		}
	}

	public void append(String value,final HSSFCellStyle style){
		try{
			if(value == null){
				value="";
			}
			cell = row.createCell(cellNo++);
			cell.setCellValue(value);
			if(style != null){
				cell.setCellStyle(style);
			}
		}catch(Exception e){

		}
	}

	public void saveFile(){
		try {
			book.write(output);
			output.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}


	public enum HeaderFooterLocation {
		LEFT,
		CENTER,
		RIGHT;
	}
}
