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
	   /* what about a Currency-class? */
		return (CLng(Expression));
	}
	/**
	 * Converts an expression into a Long.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static long CLng(VBVariant Expression) {
		/* 2009_08_03 OlimilO: must return rounded values */
		return java.lang.Math.round(Expression.doubleValue());
		//return (Expression.longValue());
	}
	/**
	 * Converts an expression into a Long.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static long CLng(Object Expression) {	   	   
		/* 2009_08_03 OlimilO: must return rounded values */
		return (CLng(CVar(Expression)));
	}
	/**
	 * Converts an expression into an Integer.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static int CInt(VBVariant Expression) {
	    /* 2009_08_03 OlimilO: must return rounded values */
		return java.lang.Math.round((float)Expression.doubleValue());
	}
	/**
	 * Converts an expression into an Integer.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static int CInt(byte Expression) {
		return (int)Expression;
	}
	
	/**
	 * Converts an expression into an Integer.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static int CInt(Object Expression) {
		return (CInt(CVar(Expression)));
	}
	/**
	 * Same like CInt
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static int Int(VBVariant Expression) {
		return CInt(Expression);
	}
	/**
	 * Same like CDbl
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static double CDbl(VBVariant Expression) {
		return (Expression.doubleValue());
	}
	/**
	 * Same like CDbl
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
	 */
	public static double CDbl(Object Expression) {
		return (CDbl(CVar(Expression)));
	}
	/**
	 * Returns the numbers contained in a string as a numeric value of appropriate type.
	 * @author Manuel Siekmann
	 * @param  Any valid numeric expression. 
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
	 * Returns the numbers contained in a string as a numeric value of appropriate type.
	 * @author Manuel Siekmann
	 * @param Append all lines in a BufferedReader and return as String
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
	 * Returns the numbers contained in a string as a numeric value of appropriate type.
	 * @author Manuel Siekmann
	 * @param Append all lines in a BufferedReader and return as String
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
	 * Concatenate all lines in a BufferedReader and return as String
	 * @author Manuel Siekmann
	 * @param Concatenate all lines in a BufferedReader and return as String
	 */
	public static String CStr(VBArrayByte Expression) {
		return (Expression.toString());
	}
	/**
	 * Concatenate all lines in a BufferedReader and return as String
	 * @author Manuel Siekmann
	 * @param Concatenate all lines in a BufferedReader and return as String
	 */
	public static String CStr(VBArrayChar Expression) {
		return (Expression.toString());
	}
	/**
	 * Returns the numbers contained in a string as a numeric value of appropriate type.
	 * @author Manuel Siekmann
	 * @param Concatenate all lines in a BufferedReader and return as String
	 */
	public static String CStr(VBVariant Expression) {
		return (Expression.toString());
	}
	/**
	 * Returns the numbers contained in a string as a numeric value of appropriate type.
	 * @author Manuel Siekmann
	 * @param Concatenate all lines in a BufferedReader and return as String
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
	public static String Str(VBVariant Expression) {
		return (CStr(Expression));
	}
/* OlimilO 2009_10_06: */
	public static VBVariant CVar(Object Expression) {
		return (new VBVariant(Expression));
	}
	public static VBVariant CVar(VBVariant Expression) {
		return (Expression);
	}
	/*public static Exception CVErr(VBVariant e) {
		return new VBVariant(e);
	}*/
	
	public static VBVariant CVErr(int Expression) {
		return new VBVariant(ErrNumberToThrowable(Expression, ""));
	}
	protected static java.lang.Throwable ErrNumberToThrowable(int errnumber, String Description) {
		switch (errnumber) {
			case 3: case 20: case 94: case 100:
				//3:   Return Ohne GoSub  //20:  Resume ohne Fehler //94:  Ungültige Verwendung von Null
				//100: Anwendungs- oder objektdefinierter Fehler
				return new java.lang.UnsupportedOperationException(Description); //InvalidOperationException(Description)
			case 5: case 446: case 448: case 449:
				//5:   Ungültiger Prozeduraurfuf oder ungültiges Argument //446: Objekt unterstützt keine benannten Argumente
				//448: Benanntes Argument nicht gefunden //449: Argument ist nicht optional
				return new java.lang.IllegalArgumentException(Description);
			case 438:
				//Objekt unterstützt diese Eigenschaft oder Methode nicht
				return new java.lang.NoSuchMethodException(Description); // MissingMemberException(Description)
			case 6:
				//6: Überlauf
				return new java.lang.StackOverflowError(Description); // StackOverflowException(Description)
			case 7: case 14:
				//7: Nicht genügend Speicher //14: Nicht genügend Zeichenfolgenspeicher
				return new java.lang.OutOfMemoryError(Description); // OutOfMemoryException(Description)
			case 9:
				//9: Index außerhalb des gültigen Bereichs
				return new java.lang.IndexOutOfBoundsException(Description); // IndexOutOfRangeException(Description)
			case 11:
				//11: Division durch Null 
				// -> is no longer an error
				return new java.lang.ArithmeticException(Description); // DivideByZeroException(Description)
			case 13:
				// Typen unverträglich
				return new java.lang.ClassCastException(Description); // InvalidCastException(Description)
			case 28:
				//28: Nicht genügend Stapelspeicher
				return new java.lang.StackOverflowError(Description); // StackOverflowException(Description)
			case 48:
				//48: Fehler beim Laden einer DLL
				return new java.lang.TypeNotPresentException(Description, null); // TypeLoadException(Description) '???
			case 53:
				//53: Datei nicht gefunden
				return new java.io.FileNotFoundException(Description); // FileNotFoundException(Description)
			case 62:
				//62: Einlesen hinter Dateiende
				return new java.io.EOFException(Description); // EndOfStreamException(Description)
			case 52: case 54: case 55: case 57: case 58: case 59: case 61: case 63: case 67: case 68: case 70: case 71: case 74: case 75:
				//52: Dateiname oder -nummer falsch
				return new java.io.IOException(Description); //
			case 76: case 432:
				//76: Pfad nicht gefunden
				return new java.io.FileNotFoundException(Description); //
			case 91:
				//91: Objektvariable oder With-Blockvariable nicht festgelegt
				return new java.lang.NullPointerException(Description); //
			case 422:
				//422: Eigenschaft nicht gefunden
				return new java.lang.NoSuchFieldException(Description); // MissingFieldException(Description)
			case 429: case 462:
			    //429: Objekterstellung durch ActiveX-Komponente nicht möglich
				//462: Der Remote-Server-Computer existiert nicht oder ist nicht verfügbar.
				return new java.lang.Exception(Description); //
			case -2147467261:
				//Automatisierungsfehler ungültiger Zeiger
				return new java.lang.RuntimeException(Description); //AccessViolationException
		}	
		return new java.lang.Exception(Description);
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
	public static VBVariant CDate(Object obj) {
		return new VBVariant(obj);
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
	    /*
		 * 2009_08_14 OlimilO: nice to have also 64bit-hex
		 */
		return (new VBVariant(Long.toHexString(Expression.longValue()).toUpperCase()));
	}
	public static String Hex(long Expression) {
		/*
		 * 2009_08_14 OlimilO: nice to have also 64bit-hex
		 */
		return java.lang.Long.toHexString(Expression).toUpperCase();
	}	
	public static String Hex(int Expression) {
		/*
		 * 2009_08_14 OlimilO: nice to have also 32bit-hex
		 */
		return java.lang.Integer.toHexString(Expression).toUpperCase();
	}	
	public static String Hex(short Expression) {
		return java.lang.Integer.toHexString(CopyB(Expression)).toUpperCase();
	}
	public static VBVariant Fix(VBVariant Expression) {
		/* 2009_08_14 OlimilO: must return only the integer part of the floatingpoint value */
		return new VBVariant(Expression.longValue());
	}
	public static VBVariant Oct(VBVariant Expression) {
		try {
			return (new VBVariant(Long.toOctalString(Expression.longValue()).toUpperCase()));
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
	public static int CopyB(short s) {
		//return (s < 0) ? (int)s ^ 0xFFFF0000: (int)s;
		//*
		if (s < 0) {
			return (int)s ^ 0xFFFF0000;	
		}
		else {
			return (int)s;
		}
		//*/
	}
	public static long CopyB(int i) {
		//return (i < 0) ? (long)i ^ 0xFFFFFFFF00000000L: (long)i;
		//*
		if (i < 0) {
			return (long)i ^ 0xFFFFFFFF00000000L;	
		}
		else {
			return (long)i;
		}
		//*/
	}
/*
'   Conversion function list
'   CBool
'   CByte
'   CCur
'   CDate
'   CDbl
'   CDec
'   CInt
'   CLng
'   CSng
'   CStr
'   CVar
'   CVDate
'   CVErr
'   Error
'   Error$
'   Fix
'   Hex
'   Hex$
'   Int
'   Oct
'   Oct$
'   Str
'   Str$
'   Val

*/

}
