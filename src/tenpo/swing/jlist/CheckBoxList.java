package tenpo.swing.jlist;

import java.awt.Component;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.JCheckBox;
import javax.swing.JList;
import javax.swing.ListCellRenderer;
import javax.swing.ListModel;
import javax.swing.ListSelectionModel;
import javax.swing.UIManager;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;

public class CheckBoxList extends JList {

	/** */
	private static final long serialVersionUID = 1L;

	protected static Border noFocusBorder = new EmptyBorder(1, 1, 1, 1);

	public CheckBoxList() {
		setCellRenderer(new CellRenderer());

		addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(final MouseEvent e) {
				int index = locationToIndex(e.getPoint());

				if (index != -1) {
					JCheckBox checkbox = (JCheckBox) getModel().getElementAt(index);
					checkbox.setSelected(!checkbox.isSelected());
					repaint();
				}
			}
		});

		setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
	}

	protected class CellRenderer implements ListCellRenderer {
		@Override
		public Component getListCellRendererComponent(final JList list,
				final Object value, final int index, final boolean isSelected,
				final boolean cellHasFocus) {
			JCheckBox checkbox = (JCheckBox) value;
			checkbox.setBackground(isSelected ? getSelectionBackground() : getBackground());
			checkbox.setForeground(isSelected ? getSelectionForeground() : getForeground());
			checkbox.setEnabled(isEnabled());
			checkbox.setFont(getFont());
			checkbox.setFocusPainted(false);
			checkbox.setBorderPainted(true);
			checkbox.setBorder(isSelected ? UIManager.getBorder("List.focusCellHighlightBorder") : noFocusBorder);
			return checkbox;
		}
	}

	public void addCheckbox(final JCheckBox checkBox) {
		ListModel currentList = this.getModel();
		JCheckBox[] newList = new JCheckBox[currentList.getSize() + 1];
		for (int i = 0; i < currentList.getSize(); i++) {
			newList[i] = (JCheckBox) currentList.getElementAt(i);
		}
		newList[newList.length - 1] = checkBox;
		setListData(newList);
	}

	public java.util.List<Integer> getCheckedIdexes() {
		java.util.List<Integer> list = new java.util.ArrayList<Integer>();
		int size = getModel().getSize();
		for (int i = 0; i < size; ++i) {
			Object obj = getModel().getElementAt(i);
			if (obj instanceof JCheckBox) {
				JCheckBox checkbox = (JCheckBox) obj;
				if (checkbox.isSelected()) {
					list.add(i);
				}
			}
		}

		return list;
	}

	public java.util.List<String> getCheckedItems() {
		java.util.List<String> list = new java.util.ArrayList<String>();
		int size = getModel().getSize();
		for (int i = 0; i < size; ++i) {
			Object obj = getModel().getElementAt(i);
			if (obj instanceof JCheckBox) {
				JCheckBox checkbox = (JCheckBox) obj;
				if (checkbox.isSelected()) {
					list.add(checkbox.getText());
				}
			}
		}
		return list;
	}
}
