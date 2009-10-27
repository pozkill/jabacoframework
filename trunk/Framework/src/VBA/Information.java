package VBA;
import java.lang.*;
import java.awt.Color;
import java.awt.SystemColor;


/*
Function Err() As ErrObject
Function IMEStatus() As VbIMEStatus
Function IsDate(Expression) As Boolean
Function IsEmpty(Expression) As Boolean
Function IsError(Expression) As Boolean
Function IsMissing(ArgName) As Boolean
Function IsObject(Expression) As Boolean
*/

public class Information {
		
	public static ErrObject EErr() {
		return ErrObject.getInstance();
	}
	/*public static void JBThrow (Exception e) throws Exception {
		throw e;
	}*/
	/*
	Converts an expression into a Boolean.
	@author Manuel Siekmann
	@param  Any valid Char or String or numeric expression. 
	*/
	public static java.awt.Color RGBtoColor(long lRGB) {
		//return new Color((int)lRGB, alpha);
		if (lRGB < 0) { 
			if (lRGB == -2147483647) return SystemColor.desktop;              // Desktop BackColor
			if (lRGB == -2147483646) return SystemColor.activeCaption;        // Window-Title BackColor (active)
			if (lRGB == -2147483639) return SystemColor.activeCaptionText;    // Window-Title Text (active)
			if (lRGB == -2147483638) return SystemColor.activeCaptionBorder;
			if (lRGB == -2147483645) return SystemColor.inactiveCaption;      // Window-Title BackColor (inactive)
			if (lRGB == -2147483629) return SystemColor.inactiveCaptionText;  // Window-Title Text (inactive)
			if (lRGB == -2147483637) return SystemColor.inactiveCaptionBorder;
			if (lRGB == -2147483643) return SystemColor.window;               
			if (lRGB == -2147483642) return SystemColor.windowBorder;         
			if (lRGB == -2147483640) return SystemColor.windowText;           
			if (lRGB == -2147483644) return SystemColor.menu;                 
			if (lRGB == -2147483641) return SystemColor.menuText;
			//if (lRGB == 0) return SystemColor.text;            // == window
			//if (lRGB == 0) return SystemColor.textText;        // == windowText
			if (lRGB == -2147483635) return SystemColor.textHighlight;        // Selection BackColor
			if (lRGB == -2147483634) return SystemColor.textHighlightText;    // Selection TextColor
			if (lRGB == -2147483633) return SystemColor.control;
			if (lRGB == -2147483630) return SystemColor.controlText;
			//if (lRGB == 0) return SystemColor.controlLtHighlight;
			//if (lRGB == 0) return SystemColor.controlHighlight;
			if (lRGB == -2147483626) return SystemColor.controlShadow;
			if (lRGB == -2147483627) return SystemColor.controlDkShadow;
			if (lRGB == -2147483631) return SystemColor.textInactiveText;
			if (lRGB == -2147483648) return SystemColor.scrollbar;
			if (lRGB == -2147483624) return SystemColor.info;
			if (lRGB == -2147483625) return SystemColor.infoText;
			return RGBtoColor(RGB(255, 55, 55)); // Mark BAD COLOR
		} else {
			return new java.awt.Color(getRfromRGB(lRGB),getGfromRGB(lRGB),getBfromRGB(lRGB));
		}
	}
	private static int getRfromRGB(long lRGB) {
		return (int)((lRGB / 0x1) & 0xff);
	}
	private static int getGfromRGB(long lRGB) {
		return (int)((lRGB / 0x100) & 0xff);
	}
	private static int getBfromRGB(long lRGB) {
		return (int)((lRGB / 0x10000) & 0xff);
	}
	private static int getAfromRGB(long lRGB) {
		return (int)((lRGB / 0x1000000) & 0xff);
	}
	public static long ColorToRGB(java.awt.Color cColor) {
		//return (BGRtoRGB(cColor.getRGB()) & 0xffffff);
		return RGB(cColor.getRed(), cColor.getGreen(), cColor.getBlue()); 
	}
	public static long BGRtoRGB(long lBGR) {
		return RGBtoBGR(lBGR);
	}
	public static long RGBtoBGR(long lRGB) {
		Color tmpColor = new Color((int)lRGB); // blue 0-7, green 8-15, red 16-23, aplha 24-31
		//return RGB(getBfromRGB(lRGB), getGfromRGB(lRGB), getRfromRGB(lRGB), getAfromRGB(lRGB));	
		return RGB(tmpColor.getRed(), tmpColor.getGreen(), tmpColor.getBlue());
	}
	public static long RGB(int R, int G, int B) {
		return RGB(R, G, B, 0);	
	}
	public static long RGB(int R, int G, int B, int A) {
		return (long)((A  * 256 * 256 * 256) + (B * 256 * 256) + (G * 256) + R);	
	}
	public static long NotRGB(long lRGB) {
		long lnotrgb = ~ lRGB;
		return (long)((0xFF000000L & lRGB) | (0xFFFFFFL & lnotrgb));
	}
	public static java.awt.Color NegativeColor(java.awt.Color col) {
		//OlimilO 27.oct.2009
		return new java.awt.Color((col.getAlpha() << 24) | (0xFFFFFF & (~ col.getRGB())), true);
	}
	public static long QBColor(int iColor) {
		switch (iColor) {
			case 0:  { return RGB(  0,  0,  0); }
			case 1:  { return RGB(  0,  0,128); }
			case 2:  { return RGB(  0,128,  0); }
			case 3:  { return RGB(  0,128,128); }
			case 4:  { return RGB(128,  0,  0); }
			case 5:  { return RGB(128,  0,128); }
			case 6:  { return RGB(128,128,  0); }
			case 7:  { return RGB(192,192,192); }
			case 8:  { return RGB(128,128,128); }
			case 9:  { return RGB(  0,  0,255); }
			case 10: { return RGB(  0,255,  0); }
			case 11: { return RGB(  0,255,255); }
			case 12: { return RGB(255,  0,  0); }
			case 13: { return RGB(255,  0,255); }
			case 14: { return RGB(255,255,  0); }
			case 15: { return RGB(255,255,255); }
		}
		return 0;
	}
	public static boolean IsFloatingPoint(double val) {
		 return (val != ((double)(long)val));
	}
	public static boolean IsNumeric(String val) {
		 try { 
			Double.parseDouble(val); 
			return true;
		} catch (Exception e){ return false; }
	}
	public static boolean IsNull(Object val) {
		 if (val == null) return true; 
		 return false;
	}
	public static boolean IsNothing(Object val) {
		 if (val == null) return true; 
		 return false;
	}
	
	public static boolean IsArray(Object val) {
		 if (val instanceof VBArray) return true;
		 return false;
	}
	public static boolean IsEnum(Object val) {
		 if (val instanceof VBEnumClass) return true;
		 return false;
	}
	public static String TypeName(Object val) {
		if (val == null) 
			return "Nothing";
		else
			return val.getClass().getName();
	}
	
	/*public static VBVarType VarType(Object val) {
		if (IsNull(val)) return VBVarType.vbNull;
		if (IsArray(val)) return VBVarType.vbArray;
		if (IsEnum(val)) return VBVarType.vbUserDefinedType;
		if (val instanceof VBVariant) return VBVarType.vbVariant;
		if (val instanceof String) { if (val == Constants.vbNullString) return VBVarType.vbEmpty; else return VBVarType.vbString;}
		if (val instanceof java.util.Date) return VBVarType.vbDate;
		if (val instanceof Exception) return VBVarType.vbError;
		if (val instanceof Integer) return VBVarType.vbInteger;
		if (val instanceof Long) return VBVarType.vbLong;
		if (val instanceof Float) return VBVarType.vbSingle;
		if (val instanceof Double) return VBVarType.vbDouble;
		if (val instanceof Boolean) return VBVarType.vbBoolean;
		if (val instanceof Byte) return VBVarType.vbByte;
		return VBVarType.vbObject;		
	}*/
	public static VBVarType VarType(VBVariant VarName) {
		return new VBVarType(VarName.VType());
	}
}
