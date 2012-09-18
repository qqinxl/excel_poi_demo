package tenpo.base.template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.security.InvalidParameterException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @author kyou
 *
 */
public class ExcelCreator {

	/** InputStream */
	private InputStream in = null;
	/** HSSFWorkbook */
	private HSSFWorkbook hssfWorkbook = null;
	/** ファイル名 */
	private String filename;

	/**
	 * @param filename　ファイル名
	 * @throws IOException　IOException
	 */
	public ExcelCreator(final String filename) throws IOException {
		File file = new File(filename);
		if (file.exists()) {
			in = new FileInputStream(file);
			hssfWorkbook = new HSSFWorkbook(in);
		} else {
			hssfWorkbook = new HSSFWorkbook();
		}

		this.filename = filename;
	}

	/**
	 * @param filename Templateファイル名
	 * @throws IOException IOException
	 */
	public void loadTemplate(final String filename) throws IOException {
		InputStream temp = new FileInputStream(filename);
		HSSFWorkbook tempBook = new HSSFWorkbook(temp);
		for(int i = 0, count = tempBook.getNumberOfSheets() ; i < count; i++) {
			HSSFSheet sheet = tempBook.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			HSSFSheet createSheet = hssfWorkbook.getSheet(sheetName);
			if (createSheet == null) {
				createSheet = hssfWorkbook.createSheet(sheetName);
			}
			//すべてコピーする
			ExcelTemplateUtils.copySheets(createSheet, sheet, true);
		}
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
			row = sheet.createRow(sheet.getLastRowNum() + 1);
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
        if (in != null) {
        	in.close();
        }
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
			newRow = sheet.createRow(sheet.getLastRowNum() + 1);
		}
		HSSFRow oldRow = sheet.getRow(oldRowNum);
		if (oldRow == null) {
			throw new InvalidParameterException();
		}
		ExcelTemplateUtils.copyRowStyle(newRow, oldRow);
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
		if (newSheet == null) {
			newSheet = hssfWorkbook.createSheet(newSheetName);
		}
		//すべてコピーする
		ExcelTemplateUtils.copySheets(newSheet, oldSheet, true);
	}

	public void deleteSheet(final String sheet) {
		HSSFSheet oldSheet = hssfWorkbook.getSheet(sheet);
		if (oldSheet == null) {
			throw new InvalidParameterException();
		}
		hssfWorkbook.removeSheetAt(hssfWorkbook.getSheetIndex(sheet));
	}

}