package userinterface.graphics;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import enums.ActivityMultiplier;
import enums.UnitSystem;

import util.GUICreator;

public class UserSettings extends JPanel {

	private static final long serialVersionUID = -1281037719857992080L;

	private JComboBox<UnitSystem> unitSystem;
	private JComboBox<ActivityMultiplier> activityLevel;
	private JPanel weight = new JPanel();
	private JPanel height = new JPanel();
	public UserSettings() {
		//super(new GridLayout(5,5));
		JLabel unitSystemL = util.GUICreator.createLabel("What unit system do you use?");
		JLabel activityLevelL = util.GUICreator.createLabel("What is your activity level?");
		
		activityLevel = new JComboBox<ActivityMultiplier>(ActivityMultiplier.values());
		activityLevel.setFont(settings.Settings.DEFAULT_FONT);
		
		unitSystem = new JComboBox<UnitSystem>(UnitSystem.values());
		unitSystem.setFont(settings.Settings.DEFAULT_FONT);
		unitSystem.addActionListener(new UnitSystemChange());
		
		//weight.setLayout(new FlowLayout());
		//height.setLayout(new FlowLayout());
		changeToUnit(UnitSystem.IMPERIAL);
		
		JButton submit = GUICreator.createButton("Submit Form", new SubmitSettings());
		this.setLayout(new FlowLayout());
		this.add(activityLevelL);
		this.add(activityLevel);
		JPanel p = new JPanel();
		p.add(unitSystemL);
		p.add(unitSystem);
		this.add(p);
		this.add(submit);
		this.add(weight);
		this.add(height);
	}
	
	public UnitSystem getUnits() { return (UnitSystem) this.unitSystem.getSelectedItem(); }
	public ActivityMultiplier getActivity() { return (ActivityMultiplier) this.activityLevel.getSelectedItem(); }
	public double getWeight() {
		
		for(Component c : this.weight.getComponents()) {
			if(c.getClass() == JTextField.class) {
				String t = ((JTextField) c).getText();
				try {
					return Double.parseDouble(t);
				}
				catch (NumberFormatException e) {
					return 0;
				}
			}
		}
		return 0;
	}
	
	public void changeToUnit(UnitSystem toUnit) {
		System.out.println("Now units are: " + toUnit);
		this.weight.removeAll();
		this.height.removeAll();
		if(toUnit.equals(UnitSystem.IMPERIAL)) {
			JLabel wLab = GUICreator.createLabel("Weight: ");
			JLabel wLab2 = GUICreator.createLabel("pounds");
			JTextField wField = new JTextField(5);
			
			JLabel hLab = GUICreator.createLabel("Height: ");
			JLabel hLab2 = GUICreator.createLabel("feet");
			JLabel hLab3 = GUICreator.createLabel("inches");
			JTextField hField1 = new JTextField(5);
			JTextField hField2 = new JTextField(5);
			
			this.weight.add(wLab);
			this.weight.add(wField);
			this.weight.add(wLab2);
			
			this.height.add(hLab);
			this.height.add(hField1);
			this.height.add(hLab2);
			this.height.add(hField2);
			this.height.add(hLab3);
		}
		else {
			JLabel wLab = GUICreator.createLabel("Weight: ");
			JTextField wField = new JTextField(5);
			//wField.addActionListener(new FieldIsDouble(this));
			JLabel wLab2 = GUICreator.createLabel("kilograms");
			
			JLabel hLab = GUICreator.createLabel("Height: ");
			JTextField hField = new JTextField(5);
			JLabel hLab2 = GUICreator.createLabel("centimeters");
			
			this.weight.add(wLab);
			this.weight.add(wField);
			this.weight.add(wLab2);
			
			this.height.add(hLab);
			this.height.add(hField);
			this.weight.add(hLab2);
		}
		this.repaint();
		this.validate();
	}
	
	public userinterface.UserSettings getSettings() {
		ActivityMultiplier al = getActivity();
		UnitSystem units = getUnits();
		
		double weight = getWeight();
		double height;
		if(units.equals(UnitSystem.IMPERIAL)) {
			
		}
		
		
		int age;
		
		userinterface.UserSettings us = new userinterface.UserSettings(250, 71, 22, UnitSystem.IMPERIAL, al);
		
		return us;
	}
	
	class FieldIsDouble implements ActionListener {
		private JTextField f;
		public FieldIsDouble(JTextField f) {
			this.f = f;
		}
		public void actionPerformed(ActionEvent e) {
			f.getText();
			
		}
	}
	
	class SubmitSettings implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			System.out.println(getSettings());
			
		}
	}
	
	class UnitSystemChange implements ActionListener {
		public void actionPerformed(ActionEvent e) {
			// this seems a bit odd
			//JComboBox<UnitSystem> box = (JComboBox<UnitSystem>) e.getSource();
			//UnitSystem units = (UnitSystem) box.getSelectedItem();
			changeToUnit(getUnits());
			//changeToUnit(units);
		}
	}

	
}
