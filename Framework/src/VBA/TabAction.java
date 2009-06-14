package VBA;

import java.awt.*;
import javax.swing.*;
import java.awt.event.*;

public class TabAction extends AbstractAction
{
	private boolean forward;

	public TabAction(boolean forward)
	{
		this.forward = forward;
	}

	public void actionPerformed(ActionEvent e)
	{
		if (forward)
			tabForward();
		else
			tabBackward();
	}

	private void tabForward()
	{
		final KeyboardFocusManager manager = KeyboardFocusManager.getCurrentKeyboardFocusManager();
		manager.focusNextComponent();

		SwingUtilities.invokeLater(new Runnable() 
		{
			public void run()
			{
				if (manager.getFocusOwner() instanceof JScrollBar)
					manager.focusNextComponent();
			}
		});
	}

	private void tabBackward()
	{
		final KeyboardFocusManager manager = KeyboardFocusManager.getCurrentKeyboardFocusManager();
		manager.focusPreviousComponent();

		SwingUtilities.invokeLater(new Runnable()
		{
			public void run()
			{
				if (manager.getFocusOwner() instanceof JScrollBar)
					manager.focusPreviousComponent();
			}
		});
	}
}
