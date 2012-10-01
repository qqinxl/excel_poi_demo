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

public class CsvCreator {

    private final static String COMMA = ",";

    public static void setCSVData(final String fileName,
            final List<LinkedHashMap<String, String>> data) throws IOException {
        if (fileName == null || data == null || data.isEmpty()) {
            throw new InvalidParameterException();
        }

        List<String> header = new ArrayList<String>(data.get(0).keySet());
        List<List<String>> listData = new ArrayList<List<String>>();

        for (LinkedHashMap<String, String> lineData : data) {
            List<String> tmpData = new ArrayList<String>();
            for (Map.Entry<String, String> e : lineData.entrySet()) {
                String value = e.getValue();
                tmpData.add(value);
            }
            listData.add(tmpData);
        }

        setCSVData(fileName, header, listData);
    }

    public static void setCSVData(final String fileName,
            final List<String> header,
            final List<List<String>> data) throws IOException {

        if (fileName == null || header == null || header.isEmpty() || data == null || data.isEmpty()) {
            throw new InvalidParameterException();
        }

        File csv = new File(fileName);
        BufferedWriter bw = new BufferedWriter(new FileWriter(csv));

        int length = header.size();

        int num = 1;
        for (String key : header) {
            bw.write(key);
            if (num < length) {
                bw.write(COMMA);
            }
            num++;
        }

        for (List<String> lineData : data) {
            bw.newLine();
            num = 1;
            for (String value : lineData) {
                bw.write(value == null ? "" : value);
                if (num < length) {
                    bw.write(COMMA);
                }
                num++;
            }
        }

        bw.close();

    }

    public static void main(final String[] args) throws IOException {
        List<LinkedHashMap<String, String>> data =
                new ArrayList<LinkedHashMap<String, String>>();

        LinkedHashMap<String, String> line1 =
                new LinkedHashMap<String, String>();
        line1.put("A11111", "11111");
        line1.put("Z22222", null);
        line1.put("B33333", "33333");
        line1.put("Y44444", "44444");
        data.add(line1);

        LinkedHashMap<String, String> line2 =
                new LinkedHashMap<String, String>();
        line2.put("A11111", "aaaaa");
        line2.put("Z22222", "zzzzz");
        line2.put("B33333", "bbbbb");
        line2.put("Y44444", "yyyyy");
        data.add(line2);

        CsvCreator.setCSVData("test.csv", data);
    }
}
