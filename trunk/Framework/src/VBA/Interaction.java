package VBA;
import java.io.*;
import java.awt.*;
import java.awt.image.*;
import javax.swing.*;
import java.io.*;

public class Interaction {
	private static Interaction MySelf = new Interaction();
	public static Object RessourceClassOwner = null;
	private static java.awt.Robot SendKeysRob = null;

	public static java.lang.reflect.Method getMethod(Object obj, String ProcName) throws Exception {
		VBA.VBArrayVariant Args = new VBA.VBArray();
		return getMethod(obj, ProcName, VBCallType.VBMethod, Args);
	}
	
	public static java.lang.reflect.Method getMethod(Object obj, String ProcName, VBCallType CallType, VBArrayVariant Args) throws Exception {
		java.lang.reflect.Method[] mets = obj.getClass().getMethods();
		Class[] metParams; 
		Class[] setParams = new Class[0]; int setParamsCount = 0; 
		if (Args.getUBound() - Args.getLBound() >= 0) {
			setParams = new Class[Args.getUBound() - Args.getLBound() + 1];
			setParamsCount = setParams.length;
			try {
				for (int i = Args.getLBound(); i <= Args.getUBound(); i++) {
					setParams[i] = (Args.valueOfVar(i)).getValClass();
				} 
			} catch(Exception e){
			}
		}
		for (int i = 0; i < mets.length; i++) {
			try{
				//if (mets[i].getDeclaringClass() == obj.getClass()){
					if (mets[i].getName().toLowerCase().equals(ProcName.toLowerCase())) {
						metParams = mets[i].getParameterTypes();
						boolean bPriority = true;
						if ( metParams.length == setParamsCount ) {
							if (setParamsCount > 0) {
								for (int ii = 0; (ii < metParams.length) && (bPriority); ii++) {
									bPriority = isPossible(metParams[ii], setParams[ii]);
								}
							}
							if (bPriority)
								return mets[i];
						}
					}
				//}
			} catch (Exception e) {
			}
		}
		return null;
	}
	
    public static boolean hasAlpha(Image image) {
        if (image instanceof BufferedImage) {
            BufferedImage bimage = (BufferedImage)image;
            return bimage.getColorModel().hasAlpha();
        }
         PixelGrabber pg = new PixelGrabber(image, 0, 0, 1, 1, false);
        try {
            pg.grabPixels();
        } catch (InterruptedException e) {
        }
        ColorModel cm = pg.getColorModel();
        return cm.hasAlpha();
    }

	public static BufferedImage createBufferedImage(Image image) {
        if (image instanceof BufferedImage) { return (BufferedImage)image; }
        image = new ImageIcon(image).getImage();
        boolean hasAlpha = hasAlpha(image);
        BufferedImage bimage = null;
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        try {
            int transparency = Transparency.OPAQUE;
            if (hasAlpha) {
                transparency = Transparency.BITMASK;
            }
            GraphicsDevice gs = ge.getDefaultScreenDevice();
            GraphicsConfiguration gc = gs.getDefaultConfiguration();
            bimage = gc.createCompatibleImage(
                image.getWidth(null), image.getHeight(null), transparency);
        } catch (HeadlessException e) {
        }
        if (bimage == null) {
            int type = BufferedImage.TYPE_INT_RGB;
            if (hasAlpha) {
                type = BufferedImage.TYPE_INT_ARGB;
            }
            bimage = new BufferedImage(image.getWidth(null), image.getHeight(null), type);
        }
        Graphics g = bimage.createGraphics();
        g.drawImage(image, 0, 0, null);
        g.dispose();
        return bimage;
    }

	
	public static String getResourceAsString(String relativeFilename) {
		File f = new File(relativeFilename);
		if (f.exists()) {
			return getResourceAsString(MySelf.getClass(), relativeFilename);
		} else {
			return getResourceAsString(MySelf.getClass(), "./" + relativeFilename);
		}
	}
	
	public static String getResourceAsString(Class relativeClass, String relativeFilename) {
		return Conversion.CStr(getResourceAsStream(relativeClass, relativeFilename));
	}
	
	public static Image getResourceAsImage(String relativeFilename) {
		//MYSELF: ./VBA/ =go=> ../ =get=> ./
		File f = new File(relativeFilename);
		if (f.exists()) {
			return getResourceAsImage(MySelf.getClass(), relativeFilename);
		} else {
			return getResourceAsImage(MySelf.getClass(), "./" + relativeFilename);
		}
	}
	
	public static Image getResourceAsImage(Class relativeClass, String relativeFilename) {
		try {
			return javax.imageio.ImageIO.read(getResourceAsStream(relativeClass, relativeFilename));
			//return Toolkit.getDefaultToolkit().createImage(getResourceAsByteArray(relativeClass, relativeFilename).toByteArray());
		} catch (Exception e) {
			return null;
		}
	}
	
	public static InputStream getResourceAsStream(Class relativeClass, String relativeFilename) {
		return relativeClass.getResourceAsStream(relativeFilename);
	}
	
	public static ByteArrayOutputStream getResourceAsByteArray(Class relativeClass, String relativeFilename) {
		byte[] myBuffer = new byte[8192];
		ByteArrayOutputStream returnValue = null;
		InputStream myInputStream = relativeClass.getResourceAsStream(relativeFilename);
		if (myInputStream != null) {
			try {
				BufferedInputStream myBufferedInput = new BufferedInputStream(myInputStream);
				ByteArrayOutputStream myBufferedOutput = new ByteArrayOutputStream(myInputStream.available());
				int count;
				while ((count = myBufferedInput.read(myBuffer, 0, myBuffer.length)) != -1) {
					myBufferedOutput.write(myBuffer, 0, count);
				}
				returnValue = myBufferedOutput;
			} catch (Exception exception) {
				return null;
			}
		} else {
			return null;
		}
		return returnValue;
	}

	public static VBVariant IIF(boolean Expression, VBVariant TruePart, VBVariant FalsePart) {
		if (Expression) {
			return TruePart;
		} else {
			return FalsePart;
		}
	}

	public static VBVariant CallByName(Object obj, String ProcName) throws Exception {
		VBA.VBArrayVariant Args = new VBA.VBArray();
		return CallByName(obj, ProcName, Args);
	}
	public static VBVariant CallByName(Object obj, String ProcName, String Arg) throws Exception {
		VBA.VBArrayVariant myArgs = new VBA.VBArray(0, 0, 0);
		myArgs.setValue(0, new VBA.VBVariant(Arg));
		return CallByName(obj, ProcName, VBCallType.VBMethod, myArgs);
	}
	public static VBVariant CallByName(Object obj, String ProcName, VBA.VBArrayVariant Args) throws Exception {
		return CallByName(obj, ProcName, VBCallType.VBMethod, Args);
	}
	
	public static VBVariant CallByName(Object obj, String ProcName, VBCallType CallType, VBA.VBArrayVariant Args) throws Exception {
		try {
			java.lang.reflect.Method Methode = getMethod(obj, ProcName, CallType, Args);
			return convertReturnParam(Methode.invoke (obj, convertClassParams(Args, Methode.getParameterTypes())));
		} catch(Exception e) {
			if (ProcName.startsWith("~") == false) {
				return CallByName(obj, "~" + ProcName, CallType, Args);
			} else {
				throw new Exception("CallByName '" + ProcName.replaceAll("~", "") + "' failed!");
			}
		}
	}
	
	public static VBVariant NativeCall(String ModuleName, String ProcName) throws Exception {
		return NativeCall(ModuleName, ProcName, (VBArrayVariant)(new VBArray()));
	}
	
	public static VBVariant NativeCall(String ModuleName, String ProcName, VBA.VBArrayVariant Args) throws Exception {
		com.eaio.nativecall.NativeCall.init();
		VBVariant ret = new VBVariant();
		Object[] myArgs = Args.objectArray();
		if (myArgs != null) {
			for (int i = 0; i < myArgs.length; i++) {
				if (myArgs[i] instanceof VBVariant) myArgs[i] = ((VBVariant)(myArgs[i])).getValObject();
				if (myArgs[i] instanceof java.lang.Long) myArgs[i] = new java.lang.Integer(((java.lang.Long)myArgs[i]).intValue());
			}
		}
		try {
			com.eaio.nativecall.IntCall tmpIntCall = new com.eaio.nativecall.IntCall(ModuleName, ProcName);
			if ((myArgs == null) || (myArgs.length == 0)) {
				ret = new VBVariant(tmpIntCall.executeCall());
			} else {
				ret = new VBVariant(tmpIntCall.executeCall(myArgs));
			}
			return ret;
		} catch(Exception e) {
			System.out.println(e.toString());
			throw new Exception("NativeCall '" + ModuleName + "." + ProcName + "' failed!");
		}
	}
	
	private static boolean isPossible(Class C1, Class C2) {
		return true;
	}
	
	private static Object[] convertClassParams(VBA.VBArrayVariant fromAry, Class[] toClassTypes) {
		Object[] ret = new Object[toClassTypes.length];
		try {
		for (int i = fromAry.getLBound(); i <= fromAry.getUBound(); i++) {
			
			if ( toClassTypes[i] == Integer.TYPE |  toClassTypes[i] == Integer.class ) {
				ret[i] = new Integer(fromAry.valueOfVar(i).intValue());
				
			} else if ( toClassTypes[i] == Boolean.TYPE | toClassTypes[i] == Boolean.class ) {
				ret[i] = new Boolean(fromAry.valueOfVar(i).intValue() == 0 ? false : true);
				
			} else if ( toClassTypes[i] == Double.TYPE | toClassTypes[i] == Double.class ) {
				ret[i] = new Double(fromAry.valueOfVar(i).doubleValue());
				
			} else if ( toClassTypes[i] == Float.TYPE |  toClassTypes[i] == Float.class ) {
				ret[i] = new Float(fromAry.valueOfVar(i).doubleValue());
				
			} else if ( toClassTypes[i] == Byte.TYPE |  toClassTypes[i] == Byte.class ) {
				ret[i] = new Byte((byte)fromAry.valueOfVar(i).intValue());
				
			} else if ( toClassTypes[i] == Long.TYPE |  toClassTypes[i] == Long.class ) {
				ret[i] = new Long(fromAry.valueOfVar(i).longValue());
				
			} else if ( toClassTypes[i] == Short.TYPE |  toClassTypes[i] == Short.class ) {
				ret[i] = new Short(((short)fromAry.valueOfVar(i).intValue()));
								
			} else if ( toClassTypes[i] == String.class ) {
				ret[i] = new String(fromAry.valueOfVar(i).toString());

			} else {
				ret[i] = fromAry.valueOfVar(i).toObject();
			}
		}
		} catch (Exception e) {}
		return ret;
	}
	private static VBVariant convertReturnParam(Object obj) {
		return new VBVariant(obj);
	}
	
	

	public static String InputBox(String Prompt) {
		return InputBox(Prompt, "", "");
	}
	public static String InputBox(String Prompt, String Title) {
		return InputBox(Prompt, Title, "");
	}
	public static String InputBox(String Prompt, String Title, String Default) {
		try {
			return (String)JOptionPane.showInputDialog(getJabacoFocusedWindow(), Prompt, Title, JOptionPane.QUESTION_MESSAGE, null, null, Default);
		} catch(Exception exception) {
			return "";
		}
	}
	private static boolean checkBitMask(long Mask, long TestBit) {
		return ((Mask & TestBit) == TestBit) ;
	}
	public static VBMsgBoxResult MsgBox(String Prompt) {
		return MsgBox (Prompt, VBMsgBoxStyle.vbOKOnly);
	}
	public static VBMsgBoxResult MsgBox(String Prompt, VBMsgBoxStyle Buttons) {
		return MsgBox (Prompt, Buttons, "");
	}
	public static VBMsgBoxResult MsgBox(String Prompt, VBMsgBoxStyle Buttons, String Title) {
		return MsgBox (getJabacoFocusedWindow(), Prompt, Buttons, Title);
	}
	private static VBMsgBoxResult MsgBox(Component Parent, String Prompt, VBMsgBoxStyle Buttons, String Title) {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName() );
		} catch( Exception e ) { e.printStackTrace(); }
		if (Title.length() == 0) Title = "Jabaco";

		int IntJavaButtons = JOptionPane.DEFAULT_OPTION; int IntJavaMessage = JOptionPane.PLAIN_MESSAGE;
		long LongButtons = 0;
		
		LongButtons = Buttons.longValue();
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbCritical.longValue())) IntJavaMessage = JOptionPane.ERROR_MESSAGE;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbQuestion.longValue())) IntJavaMessage = JOptionPane.QUESTION_MESSAGE;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbExclamation.longValue())) IntJavaMessage = JOptionPane.WARNING_MESSAGE;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbInformation.longValue())) IntJavaMessage = JOptionPane.INFORMATION_MESSAGE;

		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbOKOnly.longValue())) IntJavaButtons =  JOptionPane.DEFAULT_OPTION;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbOKCancel.longValue())) IntJavaButtons = JOptionPane.OK_CANCEL_OPTION;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbYesNoCancel.longValue())) IntJavaButtons = JOptionPane.YES_NO_CANCEL_OPTION;
		if (Interaction.checkBitMask(LongButtons, VBMsgBoxStyle.vbYesNo.longValue())) IntJavaButtons = JOptionPane.YES_NO_OPTION;
		
		int ret = JOptionPane.showConfirmDialog(Parent, Prompt, Title, IntJavaButtons, IntJavaMessage);

		VBMsgBoxResult retVB = VBMsgBoxResult.vbCancel;
		switch (ret) {
			case JOptionPane.YES_OPTION: {
				switch (IntJavaButtons) {
					case JOptionPane.YES_NO_OPTION:
					case JOptionPane.YES_NO_CANCEL_OPTION:
						retVB = VBMsgBoxResult.vbYes; 
						break;
					default:
						retVB = VBMsgBoxResult.vbOK;
						break;
				}
				break;
			}
			case JOptionPane.NO_OPTION: 	{ retVB = VBMsgBoxResult.vbNo; break; }
			case JOptionPane.CANCEL_OPTION: { retVB = VBMsgBoxResult.vbCancel; break; }
			case JOptionPane.CLOSED_OPTION: { retVB = VBMsgBoxResult.vbAbort; break; }
		}
		return retVB;
	}

	public void Beep() {
		java.awt.Toolkit.getDefaultToolkit().beep();
	}
	
	public static Component getJabacoFocusedWindow() {
		Component FocusWindow = null;
		try {
			 FocusWindow = java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager().getFocusedWindow();
			if ( FocusWindow.getLocation().getX() < -20000) {
				 FocusWindow = null; // maybe called by form_load
			}
		} catch ( Exception e ) {}
		return FocusWindow;
	}
	
	public static Frame getJabacoFocusedFrame() {
		try {
			return ((Frame)getJabacoFocusedWindow());
		} catch(Exception e) {}
		return null;
	}
	
	public static void MsgBox(Throwable ex) {
		if (ex == null) return;
        String stackTrace = "";
        try {
			StringWriter sw = new StringWriter();
			PrintWriter pw = new PrintWriter(sw);
			ex.printStackTrace(pw);
			pw.close();
			sw.close();
			stackTrace = sw.getBuffer().toString();
        } catch(Exception nex) {}
		
		ExceptionDialog exDialog = new ExceptionDialog(getJabacoFocusedFrame(), ex.toString(), stackTrace, "Jabaco Exception", true);
		exDialog.show();
	}	
	

	public static java.lang.Process Shell(String PathName) throws Exception { 
		 return Shell(PathName, VBAppWinStyle.vbNormalFocus);
	}
	
	public static java.lang.Process Shell(String PathName, VBAppWinStyle WindowStyle) throws Exception { 
		/*
			if (WindowStyle == vbHide);
			if (WindowStyle == vbNormalFocus);
			if (WindowStyle == vbMinimizedFocus);
			if (WindowStyle == vbMaximizedFocus);
			if (WindowStyle == vbNormalNoFocus);
			if (WindowStyle == vbMinimizedNoFocus);
		*/
		return Runtime.getRuntime().exec(PathName); 
	} 
	/**
	 * sends keystrokes to the current active window
	 */
	public static void SendKeys(String sKey) throws Exception {
		SendKeys(sKey, 0);
	} 
	public static void SendKeys(String sKey, boolean wait) throws Exception {
		if (wait) {
			SendKeys(sKey, -1);
		}
		else {
			SendKeys(sKey, 0);
		};
	}
	public static void SendKeys(String sKey, int wait) throws Exception {
		int a = Strings.Asc(sKey);
		java.awt.Robot r = SendKeys(a, wait);
		r = null;
	}
	public static java.awt.Robot SendKeys(int iKey) throws Exception {
		return SendKeys(iKey, 0);
	}	
	public static java.awt.Robot SendKeys(int iKey, int wait) throws Exception {
		if (SendKeysRob == null) {SendKeysRob = new java.awt.Robot();}
		if ((65<=iKey) && (iKey<=90)) {
			SendKeysRob.keyPress(java.awt.event.KeyEvent.VK_SHIFT);
			SendKeysRob.keyPress(iKey);
			SendKeysRob.keyRelease(java.awt.event.KeyEvent.VK_SHIFT);
		}
		else {
			SendKeysRob.keyPress(iKey);
		}
		if (wait > 0) {
			SendKeysRob.delay(wait);
		}
		else if (wait < 0) {
			SendKeysRob.wait();
		}
		return SendKeysRob;
	}	
	/**
	 * simulates AppActivate simply by sending [Alt]+[Tab]
	 */	
	public static void AppActivate(VBVariant Title_PID_ignored) {
        try {
			AppActivate(Title_PID_ignored, null);
		} catch(Exception e) {}
		
	}
	public static void AppActivate(VBVariant Title_PID_ignored, VBVariant wait) throws Exception {
		if (SendKeysRob == null) {SendKeysRob = new java.awt.Robot();}
		SendKeysRob.keyPress(java.awt.event.KeyEvent.VK_ALT);
		SendKeysRob.keyPress(java.awt.event.KeyEvent.VK_TAB);
		SendKeysRob.keyRelease(java.awt.event.KeyEvent.VK_TAB);
		SendKeysRob.keyRelease(java.awt.event.KeyEvent.VK_ALT);           
	}  
}
