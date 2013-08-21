package userinterface.graphics;

import java.awt.Dimension;

import javax.swing.JFrame;

public class Display extends JFrame {
	private static final int WIDTH = 850;
	private static final int HEIGHT = 680;
	private final Dimension SIZE = new Dimension(WIDTH, HEIGHT);
	
	public void display() {
		this.setTitle("Stay Fit: Excel Writer");
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setContentPane(new UserSettings());
		this.pack();
		this.setMinimumSize(SIZE);
		this.setSize(SIZE);
		this.setPreferredSize(SIZE);
		this.setResizable(false);
		this.setVisible(true);
		this.setLocationRelativeTo(null);

	}

	public static int worldWidth() { return WIDTH; }
	public static int worldHeight() { return HEIGHT; }
	
	public static void main(String[] _) {
		new Display().display();
	}
}
