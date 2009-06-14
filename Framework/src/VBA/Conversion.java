package VBA;

import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.security.*;
import java.io.*;

/**
 * The Conversion module contains the procedures used to perform various conversion operations. 
 * This module supports the Jabaco language keywords and run-time library members that convert
 * decimal numbers to other bases, numbers to strings, strings to numbers, and one data type to another.
 *
 * @author Manuel Siekmann
 * @version 1.0
 */
public class Conversion {
	private char[] alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=".toCharArray();
	private final static String[][] CheckDatePatterns = { 
			{ "yyyy-MM-dd", " hh:mm", ":ss", " a"}, 
			{ "dd.MM.yyyy", " hh:mm", ":ss", " a"}, 
			{ "dd/mm/yyyy", " hh:mm", ":ss", " a"},
			{ "#dd/mm/yyyy#", " hh:mm", ":ss", " a"},
			{ "dd-MM-yyyy", " hh:mm", ":ss", " a"}, 
			{ "E, d. MMMM", " yyyy"}, 
			{ "d. MMMM", " yyyy", " hh:mm", ":ss", " a"}, 
			{ "MMMM d", ", yyyy", " hh:mm", ":ss", " a"}, 
			{ "hh:mm", ":ss", " a"},
			{ "yyyy"}			
		};

	/**
	* Converts an expression into a Boolean.
	* @author Manuel Siekmann
	* @param  Any valid Char or String or numeric expression. 
	*/
	public static boolean CBool(VBVariant Expression) {
		if (CInt(Expression) == 0) {
			return false;
		} else {
			return true;
		}
	}
	/**
	* Converts an expression into a Byte.
	* @author Manuel Siekmann
	* @param  0 through 255 (unsigned); fractional parts are rounded.
	*/
	public static byte CByte(VBVariant Expression) {
		return ((byte)CInt(Expression));
	}
	/**
	* Converts an expression into a Byte.
	* @author Manuel Siekmann
	* @param  0 through 255 (unsigned); fractional parts are rounded.
	*/
	public static byte CByte(Object Expression) {
		return (CByte(CVar(Expression)));
	}
	/**
	* Converts an expression into a Char.
	* @author Manuel Siekmann
	* @param  0 through 65.535 (unsigned); fractional parts are rounded.
	*/
	public static char CChar(VBVariant Expression) {
		return ((char)CInt(Expression));
	}
	/**
	* Converts an expression into a Char.
	* @author Manuel Siekmann
	* @param  0 through 65.535 (unsigned); fractional parts are rounded.
	*/
	public static char CChar(Object Expression) {
		return (CChar(CVar(Expression)));
	}
	
	/**
	* Same like CLng
	* @author Manuel Siekmann
	* @param  Any valid numeric expression. 
	*/
	public static long CCur(VBVariant Expression) {
		return (CLng(Expression));
	}
	/**
	* Converts an expression into a Long.
	* @author Manuel Siekmann
	* @param  Any valid numeric expression. 
	*/
	public static long CLng(VBVariant Expression) {
		return (Expression.longValue());
	}
	/**
	* Converts an expression into a Long.
	* @author Manuel Siekmann
	* @param  Any valid numeric expression. 
	*/
	public static long CLng(Object Expression) {
		return (CLng(CVar(Expression)));
	}
	/**
	Converts an expression into an Integer.
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static int CInt(VBVariant Expression) {
		return (Expression.intValue());
	}
	/**
	Converts an expression into an Integer.
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static int CInt(byte Expression) {
		return (int)Expression;
	}
	
	/**
	Converts an expression into an Integer.
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static int CInt(Object Expression) {
		return (CInt(CVar(Expression)));
	}
	/**
	Same like CInt
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static int Int(VBVariant Expression) {
		return CInt(Expression);
	}
	/**
	Same like CDbl
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static double CDbl(VBVariant Expression) {
		return (Expression.doubleValue());
	}
	/**
	Same like CDbl
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static double CDbl(Object Expression) {
		return (CDbl(CVar(Expression)));
	}
	/**
	Returns the numbers contained in a string as a numeric value of appropriate type.
	@author Manuel Siekmann
	@param  Any valid numeric expression. 
	*/
	public static double Val(String Expression) {
	    if (Expression == null) { return 0.0; }
	    StringBuffer strBuff = new StringBuffer();
		char decimalSeparator = (new java.text.DecimalFormatSymbols()).getDecimalSeparator();
		char c; boolean bSetOperator = false; boolean bSetDecimalSeparator = false;
	    for (int i = 0; i < Expression.length() ; i++) {
	        c = Expression.charAt(i);
	        if (Character.isDigit(c)) {
	            strBuff.append(c);
	        } else if (c == '-') {
				if (bSetOperator == false) {
					bSetOperator = true;
					strBuff.append("-");
				}
			} else if (c == '.' || c == ',') {
				if (bSetDecimalSeparator == false) {
					bSetDecimalSeparator = true;
					strBuff.append(decimalSeparator);
				}
			}
	    }
		return (CDbl(new VBVariant(strBuff.toString())));
	}
	/**
	Returns the numbers contained in a string as a numeric value of appropriate type.
	@author Manuel Siekmann
	@param Append all lines in a BufferedReader and return as String
	*/
	public static String CStr(InputStream Expression) {
		try {
			byte[] tmp = new byte[Expression.available()];
			Expression.read(tmp);
			return new String(tmp);
		} catch (Exception e) {
			return "";
		}
	}
	
	/**
	Returns the numbers contained in a string as a numeric value of appropriate type.
	@author Manuel Siekmann
	@param Append all lines in a BufferedReader and return as String
	*/
	public static String CStr(BufferedReader Expression) {
		java.lang.StringBuffer ret = new java.lang.StringBuffer();
		String tmp = null;
		try {
			while ((tmp = Expression.readLine()) != null) {
				ret.append(tmp + "\n");
			}
		} catch (Exception e) {}
		return ret.toString();
	}
	/**
	Concatenate all lines in a BufferedReader and return as String
	@author Manuel Siekmann
	@param Concatenate all lines in a BufferedReader and return as String
	*/
	public static String CStr(VBArrayByte Expression) {
		return (Expression.toString());
	}
	/**
	Concatenate all lines in a BufferedReader and return as String
	@author Manuel Siekmann
	@param Concatenate all lines in a BufferedReader and return as String
	*/
	public static String CStr(VBArrayChar Expression) {
		return (Expression.toString());
	}
	/**
	Returns the numbers contained in a string as a numeric value of appropriate type.
	@author Manuel Siekmann
	@param Concatenate all lines in a BufferedReader and return as String
	*/
	public static String CStr(VBVariant Expression) {
		return (Expression.toString());
	}
	/**
	Returns the numbers contained in a string as a numeric value of appropriate type.
	@author Manuel Siekmann
	@param Concatenate all lines in a BufferedReader and return as String
	*/
	public static String CStr(Object Expression) {
		return (CStr(CVar(Expression)));
	}
	public static String CStr(char Expression) {
		char[] myExpression = new char[1];
		myExpression[0] = Expression;
		return (CStr(myExpression));
	}
	public static String CStr(char[] Expression) {
		return (new String(Expression));
	}
	public static String CStr(double Expression) {
		if (Information.IsFloatingPoint(Expression)){
			return String.valueOf(Expression);
		} else {
			return String.valueOf((long)Expression);
		}		
	}
	public static String CStr(java.util.Date Expression) {
		if (Expression == null) return (new java.util.Date()).toString();
		return (Expression.toLocaleString());
	}
	public static VBVariant CVar(Object Expression) {
		return (new VBVariant(Expression));
	}
	public static VBVariant CVar(VBVariant Expression) {
		return (Expression);
	}
	public static String Str(VBVariant Expression) {
		return (CStr(Expression));
	}
	public static java.util.Date CDate(VBVariant Expression) {
		if (Expression.isDate()) return (Expression.toDate());
		return CDate(Expression.toString());
	}
	public static java.util.Date CDate(String Expression) {
		java.util.Date ret = null;
		for (int i = 0; i < CheckDatePatterns.length; i++) {
			for (int x = 0; x < CheckDatePatterns[i].length; x++) {
				ret = CreateDateFormater(Expression, Join(CheckDatePatterns[i], CheckDatePatterns[i].length - x));
				if (ret != null) { return ret; }
			}
		}
		return new java.util.Date();
	}
	private static String Join(String[] val, int max) {
		StringBuffer strBuffer = new StringBuffer();
		for (int i = 0; i < max; i++) {
			strBuffer.append(val[i]);
		}
		return strBuffer.toString();
	}
	private static java.util.Date CreateDateFormater(String InputDate, String DatePattern) {
		try {
			SimpleDateFormat sdFormat = new SimpleDateFormat(DatePattern);
			return sdFormat.parse(InputDate);
		} catch(Exception e) {}
		return null;
	}
	public static VBVariant Hex(VBVariant Expression) {
		return (new VBVariant(Integer.toHexString(Expression.intValue())));
	}	
	public static VBVariant Fix(VBVariant Expression) {
		return new VBVariant(CInt(Expression));
	}
	public static VBVariant Oct(VBVariant Expression) {
		try {
			return (new VBVariant(Integer.toOctalString(Expression.intValue())));
		} catch(Exception e) {}
		return new VBVariant();
	}
	public static float CSng(VBVariant Expression) {
		try {
			return ((float)Expression.doubleValue());
		} catch (Exception e) {}
		return (float)0;
	}	
	public static VBVariant CDec(VBVariant Expression) {
		try {
			if (Expression.toString().toLowerCase().startsWith("&h")) {
				String val = new VBVariant(Val(Expression.toString())).toString();
				int ival = (Integer.valueOf(val, 16)).intValue();
				return new VBVariant(ival);
			}
		} catch (Exception e) {}
		return new VBVariant();
	}	
	public static java.util.Date CVDate(VBVariant Expression) {
		return CDate(Expression);
	}	



}
