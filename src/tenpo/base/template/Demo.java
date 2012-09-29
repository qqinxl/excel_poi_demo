package tenpo.base.template;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Demo {

	/**
	 * @param args
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public static void main(final String[] args) throws IOException, InvalidFormatException {
		ExcelCreator templater = new ExcelCreator("src/tenpo/base/template/tokyo.new.xls");
		templater.loadTemplate("src/tenpo/base/template/tokyo.template.xls");

		templater.importData("main", 2, 2, "YYYYYYYYYYYYYYYYYY")
				 .importData("main", 7, 2, "02-111")
			     .importData("main", 8, 2, "03-12");
		for (int i = 0; i < 5; i++) {
			templater.importData("main", 12 + i, 1, 456789)
				.importData("main", 12 + i, 2, "China")
				.importData("main", 12 + i, 3, "Buy")
				.importData("main", 12 + i, 4, "XTKS00")
				.importData("main", 12 + i, 5, 123456789)
				.importData("main", 12 + i, 6, 66)
				.importData("main", 12 + i, 7, 987654321)
				.importData("main", 12 + i, 8, "10:00:00")
				.importData("main", 12 + i, 9, "2012/02/03");
		}

		for (int i = 0; i < 4; i++) {
			templater.copyRowStyle("main", 13+i, 12);
		}

		templater.saveFile();

		templater.setBorder("main", 12, 7, ExcelCreator.Border.Bottom, HSSFCellStyle.BORDER_DOUBLE);
		templater.renameSheet("data-2", "data");
		templater.createSheet("data-4", "data-2");

		templater.saveFile();

		Desktop.getDesktop().open(new File("C:/test test test/tokyo new test.xls"));

//        try{
//            Runtime.getRuntime().exec("cmd /c start C:\\\"test test test\"\\\"tokyo new test.xls\"");
//        }catch(IOException  e){
//            e.printStackTrace();
//        }

		String filaname = "\"tokyo new test.xls\"";
        filaname = "C:\\\"test test test\"\\" + filaname;
        try{
            Runtime.getRuntime().exec("cmd /c start " + filaname);
        }catch(IOException  e){
            e.printStackTrace();
        }

//        String virefilecmd = "C:\\test test test\\viewfile.cmd";
//        String virefilecmd = "start";
//        String[] cmd = {virefilecmd, filaname};
//        try{
//        	Runtime.getRuntime().exec(cmd);
//        }catch(IOException  e){
//        	e.printStackTrace();
//        }
	}
}
