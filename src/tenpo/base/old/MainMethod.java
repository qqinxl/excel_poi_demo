package tenpo.base.old;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import tenpo.base.old.RptExcelWrapper.HeaderFooterLocation;

public class MainMethod {


	public void excute(){

		ShiozumiamRptExcelWrapper excel = new ShiozumiamRptExcelWrapper("fileName.xls", "sheet1");

		HSSFCellStyle styleCenter = excel.createStyle("styleCenter");

		excel.newLine();
		excel.append("", styleCenter);

		excel.newLine();
		excel.append("Tick Data a", styleCenter);
		excel.append("Tick Data b", styleCenter);
		excel.append("Tick Data c", styleCenter);
		excel.append("Tick Data d", styleCenter);
		excel.append("Tick Data e", styleCenter);
		excel.append("Tick Data f", styleCenter);

		excel.newLine();
		excel.newLine();
		excel.newLine();

		HSSFCellStyle styleWrapText = excel.createStyle("dddd");
		styleWrapText.setWrapText(true);

		excel.append("Tick Data 2 \nTest\nOk", styleWrapText);

		excel.mergedRegion(1, 4, 2, 5);
		excel.setLandscape(true);


		excel.setHeader("&F ......... first header\n本本本本本本           本本本本本本!!!!", HeaderFooterLocation.RIGHT);
		excel.setFooter("Page &P of &N\n&F", HeaderFooterLocation.LEFT);

		for(int i=0; i<5; i++){
			excel.append("Tick Data", styleCenter);
		}

		for(int i=100; i<105; i++){

			excel.newSheet("sheet" + i);

			excel.newLine();
			excel.append("", styleCenter);

			excel.newLine();
			excel.append("Tick Data" + i, styleCenter);

			excel.newLine();
			excel.newLine();
			excel.newLine();

			excel.append("Tick Data" + i, styleCenter);
		}


		excel.saveFile();

	}


	/**
	 * @param args
	 */
	public static void main(final String[] args) {
		System.out.println("SStart");
		// TODO Auto-generated method stub
		MainMethod method = new MainMethod();

		method.excute();
		System.out.println("EEnd");
	}



}
