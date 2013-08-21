package userinterface.graphics;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JFrame;
import javax.swing.JPanel;

public class SwitchScreen implements ActionListener {
	private String screen;
	private JPanel oldScreen;
	public SwitchScreen(JPanel oldScreen, String screen) { 
		this.oldScreen = oldScreen;
		this.screen = screen;		
	}

	public void actionPerformed(ActionEvent _) {
		Display d = (Display) oldScreen.getParent();
		JPanel gs = null;
		if(screen.equals("UserSettings")) { gs = new UserSettings(); }
		else {
			System.err.println("Invalid screen: '" + screen + "'");
		}
		d.setContentPane(gs);
		d.repaint();
		d.validate();
	}
}