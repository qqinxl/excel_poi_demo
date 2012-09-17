package tenpo.base.template;

import java.io.IOException;

public class Demo {

	/**
	 * @param args
	 * @throws IOException
	 */
	public static void main(final String[] args) throws IOException {
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

		templater.renameSheet("data-2", "data");
		templater.createSheet("data-3", "data-2");

		templater.saveFile();
	}
}
