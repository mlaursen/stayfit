package util;

import java.awt.Color;
import java.awt.Font;
import java.awt.LayoutManager;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.border.Border;


public class GUICreator {

	public static JButton createButton(String bName, Font font,	ActionListener listener) {
		JButton b = new JButton(bName);
		b.setFont(font);
		b.addActionListener(listener);
		return b;
	}
	
	public static JButton createButton(String bName, ActionListener listener) {
		return createButton(bName, settings.Settings.DEFAULT_FONT, listener);
	}

	public static JLabel createLabel(String name, Font f, Color c) {
		JLabel l = new JLabel(name);
		l.setFont(f);
		l.setForeground(c);
		return l;
	}
	
	public static JLabel createLabel(String name) {
		return createLabel(name, settings.Settings.DEFAULT_FONT, settings.Settings.DEFAULT_COLOR);
	}
	
	public static JPanel createPanel(LayoutManager l, Border b) {
		JPanel p = new JPanel();
		p.setLayout(l);
		p.setBorder(b);
		return p;
	}
	
	public static JPanel createPanel(LayoutManager l, Border b, boolean opaque) {
		JPanel p = new JPanel();
		p.setLayout(l);
		p.setBorder(b);
		p.setOpaque(opaque);
		return p;
	}
}
