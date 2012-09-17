package org.shin.demo.excel;

import org.shin.demo.base.PropertiesLoader;

public class ExcelDemo {
	public static void exportData() {
		System.out.println("test");
//		PropertiesLoader.init();
		String value = PropertiesLoader.getAreaValue(13);
		System.out.println(value);
	}
}
