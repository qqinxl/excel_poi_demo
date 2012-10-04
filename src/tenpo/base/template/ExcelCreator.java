package tenpo.base.template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.security.InvalidParameterException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author kyou
 *
 */
public class ExcelCreator {

	/** HSSFWorkbook */
	private HSSFWorkbook hssfWorkbook = null;
	/** ファイル名 */
	private String filename;

	/**
	 * @param filename　ファイル名
	 * @throws IOException　IOException
	 * @throws InvalidFormatException
	 */
	public ExcelCreator(final String filename) throws IOException, InvalidFormatException {
		File file = new File(filename);
		if (!file.exists()) {
			hssfWorkbook = new HSSFWorkbook();
			OutputStream out = new FileOutputStream(filename);
			hssfWorkbook.write(out);
	        out.close();
	        file = new File(filename);
		}

		FileInputStream in = new FileInputStream(file);
		hssfWorkbook = (HSSFWorkbook) WorkbookFactory.create(in);
		in.close();

		this.filename = filename;
	}

	/**
	 * @param filename Templateファイル名
	 * @throws IOException IOException
	 * @throws InvalidFormatException
	 */
	public void loadTemplate(final String filename) throws IOException, InvalidFormatException {
		FileInputStream temp = new FileInputStream(filename);
		hssfWorkbook = (HSSFWorkbook) WorkbookFactory.create(temp);

		FileOutputStream out = new FileOutputStream(filename);
		hssfWorkbook.write(out);
        out.close();
	}

	/**
	 * @param sheetName　シート名。シートが存在してない場合、Exceptionを発生する。
	 * @param data　投入するデータ
	 * @param rowNum 行
	 * @param cellNum 列
	 * @param value　入力データ
	 */
	public ExcelCreator importData(final String sheetName, final int rowNum, final int cellNum, final Object value) {
		HSSFSheet sheet = hssfWorkbook.getSheet(sheetName);
		if (sheet == null) {
			throw new InvalidParameterException();
		}

		HSSFRow row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		HSSFCell cell = row.getCell(cellNum);
		if (cell == null) {
			cell = row.createCell(cellNum);
		}

		if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if (value instanceof Double) {
			cell.setCellValue((Double) value);
		} else if (value instanceof BigDecimal) {
			cell.setCellValue(((BigDecimal) value).doubleValue());
		} else if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Long) {
			cell.setCellValue((Long) value);
		} else if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else {
			cell.setCellValue(value.toString());
		}
		return this;
	}

	/**
	 * 保存する
	 * @throws IOException　IOException
	 */
	public void saveFile() throws IOException {
		OutputStream out = new FileOutputStream(filename);
		hssfWorkbook.write(out);
        out.close();
	}

	/**
	 * Styleをコピーする
	 * @param sheetName シート名。シートが存在してない場合、Exceptionを発生する。
	 * @param newRowNum 目標シート
	 * @param oldRowNum コピー元シート
	 */
	public void copyRowStyle(final String sheetName, final int newRowNum, final int oldRowNum) {
		HSSFSheet sheet = hssfWorkbook.getSheet(sheetName);
		if (sheet == null) {
			throw new InvalidParameterException();
		}

		HSSFRow newRow = sheet.getRow(newRowNum);
		if (newRow == null) {
			newRow = sheet.createRow(newRowNum);
		}
		HSSFRow oldRow = sheet.getRow(oldRowNum);
		if (oldRow == null) {
			throw new InvalidParameterException();
		}
		copyRowStyle(newRow, oldRow);
	}

	/**
	 * Styleをコピーする
	 * @param sheetName シート名。シートが存在してない場合、Exceptionを発生する。
	 * @param newRowNum 目標シート
	 * @param oldSheetNameコピー元シート
	 * @param oldRowNum コピー元シート
	 */
	public void copyRowStyle(final String sheetName, final int newRowNum, final String oldSheetName, final int oldRowNum) {
		HSSFSheet sheet = hssfWorkbook.getSheet(sheetName);
		HSSFSheet oldSheet = hssfWorkbook.getSheet(oldSheetName);
		if (sheet == null) {
			throw new InvalidParameterException();
		}
		if (oldSheet == null) {
			throw new InvalidParameterException();
		}

		HSSFRow newRow = sheet.getRow(newRowNum);
		if (newRow == null) {
			newRow = sheet.createRow(sheet.getLastRowNum() + 1);
		}
		HSSFRow oldRow = oldSheet.getRow(oldRowNum);
		if (oldRow == null) {
			throw new InvalidParameterException();
		}
		copyRowStyle(newRow, oldRow);
	}

	/**
	 * シートをRenameする
	 * @param newSheetName
	 * @param oldSheetName
	 */
	public void renameSheet(final String newSheetName, final String oldSheetName) {
		HSSFSheet newSheet = hssfWorkbook.getSheet(newSheetName);
		if (newSheet != null) {
			return;
		}

		int index = hssfWorkbook.getSheetIndex(oldSheetName);
		hssfWorkbook.setSheetName(index, newSheetName);
	}

	/**
	 * 新規シートを追加する
	 * @param newSheetName
	 * @param templateSheetName 模板シート
	 */
	public void createSheet(final String newSheetName, final String templateSheetName) {
		HSSFSheet oldSheet = hssfWorkbook.getSheet(templateSheetName);
		if (oldSheet == null) {
			throw new InvalidParameterException();
		}
		HSSFSheet newSheet = hssfWorkbook.getSheet(newSheetName);
		if (newSheet != null) {
			return;
		}
		//すべてコピーする
		HSSFSheet newSheet2 = hssfWorkbook.cloneSheet(hssfWorkbook.getSheetIndex(templateSheetName));
		renameSheet(newSheetName, newSheet2.getSheetName());
	}

	public void deleteSheet(final String sheetName) {
		HSSFSheet oldSheet = hssfWorkbook.getSheet(sheetName);
		if (oldSheet == null) {
			throw new InvalidParameterException();
		}
		hssfWorkbook.removeSheetAt(hssfWorkbook.getSheetIndex(sheetName));
	}


	// templater.setBorder("main", 12, 7, ExcelCreator.Border.Bottom, HSSFCellStyle.BORDER_DOUBLE);
	public void setBorder(final String sheetName, final int rowNum, final int cellNum, final Border border, final short borderSize) {
		if (border == null) {
			throw new InvalidParameterException();
		}
		HSSFSheet sheet = hssfWorkbook.getSheet(sheetName);
		if (sheet == null) {
			throw new InvalidParameterException();
		}

		HSSFRow row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		HSSFCell cell = row.getCell(cellNum);
		if (cell == null) {
			cell = row.createCell(cellNum);
		}
		HSSFCellStyle cellStyle = cell.getCellStyle();
		if (border.equals(Border.TOP)) {
			cellStyle.setBorderTop(borderSize);
		} else if (border.equals(Border.Bottom)) {
			cellStyle.setBorderBottom(borderSize);
		} else if (border.equals(Border.Left)) {
			cellStyle.setBorderLeft(borderSize);
		} else if (border.equals(Border.Right)) {
			cellStyle.setBorderRight(borderSize);
		}
	}
	 /**
	  * @param oldRow コピー元シート
	  * @param newRow 目標シート
	  */
	 private void copyRowStyle(final HSSFRow newRow, final HSSFRow oldRow) {
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
	 private void copyCellStyle(final HSSFCell newCell, final HSSFCell oldCell) {
        HSSFCellStyle newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
        newCell.setCellStyle(newCellStyle);
	 }

	 public void copyCellStyle(final String sheetName, final int newY, final int newX, final int oldY, final int oldX) {
		HSSFSheet sheet = hssfWorkbook.getSheet(sheetName);
		if (sheet == null) {
			throw new InvalidParameterException();
		}
		HSSFRow newRow = sheet.getRow(newY);
		if (newRow == null) {
			throw new InvalidParameterException();
		}
		HSSFCell newCell = newRow.getCell(newX);
		if (newCell == null) {
			throw new InvalidParameterException();
		}
		HSSFRow oldRow = sheet.getRow(oldY);
		if (oldRow == null) {
			throw new InvalidParameterException();
		}
		HSSFCell oldCell = oldRow.getCell(oldX);
		if (oldCell == null) {
			throw new InvalidParameterException();
		}

		copyCellStyle(newCell, oldCell);
	 }

	 public static enum Border {
		 TOP,
		 Bottom,
		 Right,
		 Left;
	 }
}