package tenpo.base.template;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

public abstract class AbstractCltReport implements cltReportIF {

	@Override
	public abstract void createReport();

	@Override
	public abstract ArrayList<String> getReportFiles();

	protected DecimalFormat formater = new DecimalFormat("#,###");

	protected SimpleDateFormat yyyyMMddDateformatter = new SimpleDateFormat("yyyyMMdd");

}
