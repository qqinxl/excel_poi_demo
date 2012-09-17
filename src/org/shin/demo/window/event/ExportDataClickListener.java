package org.shin.demo.window.event;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import org.shin.demo.excel.ExcelDemo;

public class ExportDataClickListener implements ActionListener {

	@Override
	public void actionPerformed(final ActionEvent e) {
		// TODO Auto-generated method stub
		ExcelDemo.exportData();
	}

}
