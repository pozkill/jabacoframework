package VBA;

import java.awt.*;
import sun.awt.*;


public class JabacoEventQueque extends EventQueue {

	private static JabacoEventQueque mySelf = null;

	public static JabacoEventQueque getJabacoEventQueque() {
		if (mySelf == null) {
			mySelf = new JabacoEventQueque();
			Toolkit.getDefaultToolkit().getSystemEventQueue().push(mySelf);
		}
		return mySelf;
	}

	public void dispatchEvent(AWTEvent evt) {
		super.dispatchEvent(evt);
	}

	public void dispatchEvents() throws InterruptedException {
		if (EventQueue.isDispatchThread()) {
			AWTEvent evt = null;
			while ((evt = getNonBlockingNextEvent()) != null) {
				dispatchEvent(evt);
			}
		} else {
			Thread.yield();
		}
	}
	
	public void flushPendingEvents() {
		try {
			SunToolkit.flushPendingEvents();
		} catch (Exception ex1) {
			try {
				postEvent(null); // flush indirect (for security exceptions)
			} catch (Exception ex2) {
				// ignore
			}
		}
	}

	public AWTEvent getNonBlockingNextEvent() throws InterruptedException {
		flushPendingEvents();
		if(super.peekEvent() != null)
			return super.getNextEvent();
		return null;
	}



}