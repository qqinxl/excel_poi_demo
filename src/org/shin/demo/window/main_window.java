package org.shin.demo.window;

import java.awt.EventQueue;

import javax.swing.JButton;
import javax.swing.JFrame;

import org.shin.demo.window.event.ExportDataClickListener;

public class main_window {

	private JFrame frame;

	/**
	 * Launch the application.
	 */
	public static void main(final String[] args) {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				try {
					main_window window = new main_window();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public main_window() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 517, 469);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		JButton btnNewButton = new JButton("Export Data");
		btnNewButton.addActionListener(new ExportDataClickListener());
		btnNewButton.setBounds(128, 166, 112, 58);
		frame.getContentPane().add(btnNewButton);
	}

}
