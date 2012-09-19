package tenpo.base.template;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class Shiooooo extends AbstractCltReport {


	private String yyyyMMdd = "";

	public Shiooooo () {
		Date now = new Date();
		yyyyMMdd = yyyyMMddDateformatter.format(now);
	}

	@Override
	public void createReport() {
		// TODO Auto-generated method stub

		ArrayList<ArrayList<String>> oldData = new ArrayList<ArrayList<String>>();
		Map<String, ArrayList<ArrayList<String>>> dataGroubyAccountName = new HashMap<String, ArrayList<ArrayList<String>>>();

		for (ArrayList<String> record : oldData) {
			String accountName = ""; //
			if (dataGroubyAccountName.containsKey(accountName)) {
				dataGroubyAccountName.get(accountName).add(record);
			} else {
				ArrayList<ArrayList<String>> tmp = new ArrayList<ArrayList<String>>();
				tmp.add(record);
				dataGroubyAccountName.put(accountName, tmp);
			}
		}

		for (Map.Entry<String, ArrayList<ArrayList<String>>> e : dataGroubyAccountName.entrySet()) {
			String accountName = e.getKey();
			ArrayList<ArrayList<String>> valueList = e.getValue();


		}

	}

	@Override
	public ArrayList<String> getReportFiles() {
		// TODO Auto-generated method stub
		return null;
	}

}
