package tenpo.swing.jlist;

import java.util.List;


public class CitSumitomoAM {

	/**
	 * @param args
	 */
	public static void main(final String[] args) {
		// TODO Auto-generated method stub
		CitSumitomoAM citSumitomoAM = new CitSumitomoAM();
		citSumitomoAM.createReport();
	}

	public void createReport() {
		System.out.println("Start to open a window.");
		GuiXXXSelector guiXXXSelector = new GuiXXXSelector(this);
		guiXXXSelector.setVisible(true);
	}

	public void show(final List<String> selectedList01,  final List<String> selectedList02) {
		System.out.println(selectedList01);
		System.out.println(selectedList02);

		//TODO
	}
}
