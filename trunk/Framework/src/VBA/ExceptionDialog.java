package VBA;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import javax.swing.border.*;
import java.lang.reflect.Method; 
import java.util.Arrays; 
import javax.swing.JOptionPane; 
import java.lang.*;


public class ExceptionDialog extends JDialog implements ActionListener {
	boolean bShowDetails = true;
	boolean bCreateLAF = setDefaultUI();
	public JPanel content = new JPanel();
	JButton btnStack = new JButton("---");
	JButton btnOK = new JButton("<html><u>C</u>ontinue</html>");
	JButton btnQuit = new JButton("<html><u>Q</u>uit</html>");

	public void actionPerformed(ActionEvent e) { 
		if (e.getSource() == btnStack) {
			showDetails(!bShowDetails); return;
		}	
		if (e.getSource() == btnOK) {
			dispose(); return;
		}
		if (e.getSource() == btnQuit) {
			System.exit(0); return;
		}
	}

	public ExceptionDialog(Frame owner, String message, String stackTrace, String title, boolean modal) {
		super(owner, title, modal);

		getContentPane().add(content);
		content.setBackground(getContentPane().getBackground());
		content.setLayout(null);
		// CREATE ICON
		JLabel icon = new JLabel(UIManager.getIcon("OptionPane.errorIcon"));//new ImageIcon("icon.gif"));
		icon.setSize(64, 64);
		icon.setLocation(0, 0);
		icon.setBackground(Color.orange);
		content.add(icon);

		// CREATE MESSAGE TEXT
		JTextArea txtDefaultMessage = new JTextArea("An unhandled exception has occurred in your application. If you click Continue, the application will ignore this error and attempt to continue. If you click Quit, the application will be shut down.");
		txtDefaultMessage.setFont(new Font("Arial", Font.PLAIN, 11));
		txtDefaultMessage.setBorder(null);
		txtDefaultMessage.setSize(360, 50);
		txtDefaultMessage.setLocation(66, 5);
		txtDefaultMessage.setBackground(content.getBackground());
		txtDefaultMessage.setEditable(false);
		txtDefaultMessage.setLineWrap(true);	
		txtDefaultMessage.setWrapStyleWord(true);		
		content.add(txtDefaultMessage);
		
		if (message == null) message = "unknown";
		JTextArea txtMessage = new JTextArea(message);
		txtMessage.setFont(new Font("Arial", Font.BOLD, 11));
		txtMessage.setBorder(null);
		txtMessage.setSize(360, 60);
		txtMessage.setLocation(66, 60);
		txtMessage.setBackground(content.getBackground());
		txtMessage.setEditable(false);
		txtMessage.setLineWrap(true);
		txtMessage.setWrapStyleWord(true);
		content.add(txtMessage);

		// CREATE BUTTONS
		JPanel ButtonArea = new JPanel();
		ButtonArea.setSize(440, 28);
		ButtonArea.setLocation(0, 122);
		ButtonArea.setBackground(content.getBackground());
		ButtonArea.setLayout(null);
		
		btnStack.setMnemonic(KeyEvent.VK_D);
		btnStack.setSize(110, 23);
		btnStack.setLocation(8, 2);
		btnStack.setFont(new Font("Arial", Font.PLAIN, 11));
		//ActionListener lstStack = new ActionListener() { public void actionPerformed(ActionEvent e) { showDetails(!bShowDetails); } };
		btnStack.addActionListener(this);
		ButtonArea.add(btnStack);

		btnOK.setMnemonic(KeyEvent.VK_C);
		btnOK.setSize(110, 23);
		btnOK.setLocation(200, 2);
		btnOK.setFont(new Font("Arial", Font.PLAIN, 11));
		//ActionListener lstOK = new ActionListener() { public void actionPerformed(ActionEvent e) { dispose(); } };
		btnOK.addActionListener(this);
		ButtonArea.add(btnOK);
		
		btnQuit.setMnemonic(KeyEvent.VK_Q);
		btnQuit.setSize(110, 23);
		btnQuit.setLocation(315, 2);
		btnQuit.setFont(new Font("Arial", Font.PLAIN, 11));
		//ActionListener lstQuit = new ActionListener() { public void actionPerformed(ActionEvent e) { System.exit(0); } };
		btnQuit.addActionListener(this);
		ButtonArea.add(btnQuit);
		content.add(ButtonArea);
		
		// CREATE STACK TRACE TEXT
		if (stackTrace == null) stackTrace = "";
		JTextArea txtStackTrace = new JTextArea(stackTrace);
		JScrollPane scpStackTrace = new JScrollPane(txtStackTrace, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scpStackTrace.setSize(418, 140);
		scpStackTrace.setLocation(8, 160);
		txtStackTrace.setFont(new Font("Arial", Font.PLAIN, 11));
		txtStackTrace.setEditable(false);
		txtStackTrace.setBackground(content.getBackground());
		content.add(scpStackTrace);
		
		setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		showDetails(bShowDetails);
		// set defaults
		//setSize(400, new Double(220 + LabelArea.getSize().getHeight()).intValue());		
		//setSize(400, 250);		
		setResizable(false);
		centerMySelf();
	}
	
	private void showDetails(boolean v) {
		bShowDetails = v;
		if (v) {
			setSize(440, 340);
			btnStack.setLabel("<html>&lt;&lt;&lt; <u>D</u>etails</html>");
		} else {
			setSize(440, 182); 
			btnStack.setLabel("<html><u>D</u>etails &gt;&gt;&gt;</html>");
		}
	}
	
	private boolean setDefaultUI() {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName() );
		} catch( Exception e ) { e.printStackTrace(); }
		return true;
	}
	
	private void centerMySelf() {
		Toolkit toolkit = Toolkit.getDefaultToolkit();  
		Dimension screenSize = toolkit.getScreenSize();  
		int x = (screenSize.width - getWidth()) / 2;  
		int y = (screenSize.height - getHeight()) / 2;  
		setLocation(x, y); 
	}
/*
	public static void main(String[] args) {
		String message = "message";
		(new ExceptionDialog(null, "java.lang.IndexOutOfBounds", message, "Jabaco Exception", true)).setVisible(true);
	}
*/

}

