package tenpo.base.template;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.security.InvalidParameterException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class CsvCreator {

	private final static String COMMA = ",";

	public static void setCSVData(final String fileName, final List<LinkedHashMap<String, String>> data) throws IOException {
		if (fileName == null || data == null || data.isEmpty()) {
			throw new InvalidParameterException();
		}

		File csv = new File(fileName);
		BufferedWriter bw = new BufferedWriter(new FileWriter(csv));

		Set<String> keySet = data.get(0).keySet();
		int length = keySet.size();

		int num = 1;
		for (String key : keySet) {
			bw.write(key);
			if (num < length) {
				bw.write(COMMA);
			}
			num++;
		}

		for (LinkedHashMap<String, String> lineData : data) {
			bw.newLine();
			num = 1;
			for (Map.Entry<String, String> e : lineData.entrySet()) {
				String value = e.getValue();
				bw.write(value == null ? "": value);
				if (num < length) {
					bw.write(COMMA);
				}
				num++;
			}
		}

		bw.close();

	}

	public static void main(final String[] args) throws IOException {
		List<LinkedHashMap<String, String>> data = new ArrayList<LinkedHashMap<String, String>>();

		LinkedHashMap<String, String> line1 = new LinkedHashMap<String, String>();
		line1.put("A11111", "11111");
		line1.put("Z22222", null);
		line1.put("B33333", "33333");
		line1.put("Y44444", "44444");
		data.add(line1);

		LinkedHashMap<String, String> line2 = new LinkedHashMap<String, String>();
		line2.put("A11111", "aaaaa");
		line2.put("Z22222", "zzzzz");
		line2.put("B33333", "bbbbb");
		line2.put("Y44444", "yyyyy");
		data.add(line2);

		setCSVData("test.csv", data);
	}
}
