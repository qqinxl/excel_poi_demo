package tenpo.swing.jlist;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;

public class GuiXXXSelector extends JFrame {

	/** */
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private CheckBoxList list01;
	private CheckBoxList list02;
	private List<String> selectedList01;
	private List<String> selectedList02;
	private CitSumitomoAM citSumitomoAM;
	/**
	 * Create the frame.
	 */
	public GuiXXXSelector(final CitSumitomoAM citSumitomoAM) {

		this.citSumitomoAM = citSumitomoAM;

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 774, 321);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		list01 = new CheckBoxList();
		list01.setBounds(25, 109, 179, 146);
		contentPane.add(list01);
		String[] initData01 = { "Blue", "Green", "Red", "Whit", "Black" };
		for (String key : initData01) {
			list01.addCheckbox(new JCheckBox(key));
		}

		list02 = new CheckBoxList();
		list02.setBounds(243, 109, 369, 146);
		contentPane.add(list02);
		String[] initData02 = { "俺", "私", "僕" };
		for (String key : initData02) {
			list02.addCheckbox(new JCheckBox(key));
		}

		JButton btnNewButton = new JButton("Ok");
		btnNewButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent arg0) {
				selectedList01 = list01.getCheckedItems();
				selectedList02 = list02.getCheckedItems();

				System.out.println("Window Close!");
				citSumitomoAM.show(selectedList01, selectedList02);
				dispose();
			}
		});
		btnNewButton.setBounds(647, 106, 91, 21);
		contentPane.add(btnNewButton);

		JLabel lblNewLabel = new JLabel("Please select xxxxxx");
		lblNewLabel.setBounds(25, 64, 257, 35);
		contentPane.add(lblNewLabel);

		JLabel label = new JLabel("Please select xxxxxx");
		label.setBounds(243, 64, 361, 35);
		contentPane.add(label);

		JLabel lblXxxxxxSelector = new JLabel("xxxxxx Selector");
		lblXxxxxxSelector.setHorizontalAlignment(SwingConstants.CENTER);
		lblXxxxxxSelector.setFont(new Font("MS UI Gothic", Font.PLAIN, 24));
		lblXxxxxxSelector.setBounds(25, 10, 595, 44);
		contentPane.add(lblXxxxxxSelector);
	}
}
