package org.shin.demo.base;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class PropertiesLoader {
//	private static String configFile = "src/test.properties";
	private static String configFile = "test.properties";
	private static Properties configProperties = null;

	private static String AREA = "area";
	private static String DOT = ".";

	public static void init() {
		if (configProperties == null) {
			configProperties = new Properties();
			InputStream in = null;
			try {
//				in = new FileInputStream(configFile);
				in = PropertiesLoader.class.getClassLoader().getResourceAsStream(configFile);
				configProperties.load(in);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} finally {
				if (in != null) {
					try {
						in.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		}
	}

	/**
	 * @param id
	 * @return
	 */
	public static String getAreaValue(final int id) {
		init();

		String key = AREA + DOT + String.valueOf(id);
//		String key = "area." + String.valueOf(id);
		String value = configProperties.getProperty(key);
		return value;

	}

	public static void main(final String[] args) {
		String value = PropertiesLoader.getAreaValue(13);
		System.out.println(value);
	}
}
