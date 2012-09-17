package tenpo.base.old;

import org.apache.poi.hssf.usermodel.HSSFHeader;

public class ShiozumiamRptExcelWrapper extends RptExcelWrapper {

	public ShiozumiamRptExcelWrapper(final String fileName, final String sheetName) {
		super(fileName, sheetName);
	}

	// 0:Left; 1:center; 2:right
	@Override
	public void setHeader(final String msg, final HeaderFooterLocation location) {
		String styleMsg = HSSFHeader.startDoubleUnderline() + HSSFHeader.font("Stencil-Normal", "Bold") + HSSFHeader.fontSize((short) 14) + msg + HSSFHeader.endDoubleUnderline();
		super.setHeader(styleMsg, location);
	}
	@Override
	public void setFooter(final String msg, final HeaderFooterLocation location) {
		String styleMsg = HSSFHeader.startUnderline() + HSSFHeader.font("Stencil-Normal", "Bold") + HSSFHeader.fontSize((short) 16) + msg + HSSFHeader.endUnderline();
		super.setFooter(styleMsg, location);
	}
}
