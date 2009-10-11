package VBA;

import java.awt.*;
import sun.awt.*;


public class JabacoEventQueue extends EventQueue {

	private static JabacoEventQueue mySelf = null;

	public static JabacoEventQueue getJabacoEventQueue() {
		if (mySelf == null) {
			mySelf = new JabacoEventQueue();
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