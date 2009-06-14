package VBA;

import java.awt.event.*;

public class VBMouseEvent {
	
	public static int getVBMouseButton(MouseEvent e) {
		return (
				  (((e.getModifiers() & MouseEvent.BUTTON1_MASK) != 0) ? 1 : 0) 
				| (((e.getModifiers() & MouseEvent.BUTTON3_MASK) != 0) ? 2 : 0) 
				| (((e.getModifiers() & MouseEvent.BUTTON2_MASK) != 0) ? 4 : 0) 
		       ); 
	}
	
	public static int getVBMouseShift(MouseEvent e) {
		/*	strg - 2 alt = 4 shift = 1 alt gr = 6*/
		return ( 
				 (e.isShiftDown()   ? 1 : 0)
			   | (e.isControlDown() ? 2 : 0) 
			   | (e.isAltDown()     ? 4 : 0) 
			  ); 
	}
}
