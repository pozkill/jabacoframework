package VBA;
import java.lang.StringBuffer;
/**
 * Prozeduren, die zum Durchführen von Zeichenfolgenoperationen verwendet werden
 *
 * @author Manuel Siekmann
 * @version 1.0
 */
public class Strings {
	/*
	 *@param name description
	 *@return description
	*/
	public static int AscB(String String1) {
		return Asc(String1);
	}
	public static int AscW(String String1) {
		return Asc(String1);
	}
	public static int Asc(String String1) {
		return (int)String1.charAt(0);
	}
	public static String ChrB(int Char) {
		return Chr(Char);
	}
	public static String ChrW(int Char) {
		return Chr(Char);
	}
	public static String Chr(int Char) {
		byte[] b = { (byte)Char };
		try {
			return (new String(b, "ISO-8859-1"));
		} catch (Exception e) {
			return (new String(b));
		}
	}
	public static int InStr(String String1, String String2) {
		return InStr(1, String1, String2);
	}
	public static int InStr(int Start, String String1, String String2) {
		return (InStr(Start, String1, String2, VBCompareMethod.vbBinaryCompare));
	}
	public static int InStr(int Start, String String1, String String2, VBCompareMethod CompareMet) {
		if (String1.equals("")) return 0;
		if (CompareMet.intValue() == 1) {
			String1 = String1.toLowerCase();
			String2 = String2.toLowerCase();
		}
		return String1.indexOf (String2, Start - 1) + 1;
	}
	public static int InStrRev(String String1, String String2) {
		return (InStrRev(String1, String2, Len(String1)));
	}
	public static int InStrRev(String String1, String String2, int Start) {
		return (InStrRev(String1, String2, Start, VBCompareMethod.vbBinaryCompare));
	}
	public static int InStrRev(String String1, String String2, int Start, VBCompareMethod CompareMet) {
		if (Start == -1) Start = String1.length();
		if (String1.length() > Start) String1=String1.substring(0,Start);
		if (String1.equals("") || (Start > Len(String1)) ) return 0;
		if (CompareMet.intValue() == 1) {
			String1 = String1.toLowerCase();
			String2 = String2.toLowerCase();
		}
		return String1.lastIndexOf (String2, Start - 1) + 1;
	}
	public static int Len(String String1) {
		return String1.length();
	}
	public static String Mid(String String1, int Start) {
		return Mid(String1, Start, String1.length() - Start + 1);	
	}
	public static String Mid(String String1, int Start, int Length) {
		if (Start <= 0) return Constants.vbNullString;
		int iStartPos = Start - 1;
		int iEndPos = iStartPos + Length;
		if (iStartPos > String1.length()) return "";
		if (iEndPos > String1.length()) iEndPos = String1.length();
		return String1.substring(iStartPos, iEndPos);	
	}
	public static String Right(String String1, int Length) {
		int iStartPos = String1.length()-Length;
		if (iStartPos < 0) iStartPos = 0;
		return String1.substring(iStartPos, String1.length());	
	}
	public static String Left(String String1, int Length) {
		int iEndPos = Length;
		if (iEndPos > String1.length()) iEndPos = String1.length();
		return String1.substring(0, iEndPos);	
	}
	public static String StrMerge(String String1, String String2) {
		if (Information.IsNumeric(String1) && Information.IsNumeric(String2)) {
			return Conversion.CStr(Conversion.CDbl(String1) + Conversion.CDbl(String2));
		} else {
			return StrCat(String1, String2);
		}		
	}	
	public static String StrCat(String String1, String String2) {
		StringBuffer myStringBuffer = new StringBuffer();
		myStringBuffer.append(String1);
		myStringBuffer.append(String2);
		return myStringBuffer.toString();
	}
	public static String Replace(String Source, String Pattern, String Replace) {
		final int len = Pattern.length();
		if (len == 0) return Source;
		StringBuffer sb = new StringBuffer();
		int found = -1;
		int start = 0;
		while( (found = Source.indexOf(Pattern, start) ) != -1) {
			sb.append(Source.substring(start, found));
			sb.append(Replace);
			start = found + len;
		}
		sb.append(Source.substring(start));
		return sb.toString();
	}

	public static String LCase(String String1) {
		return String1.toLowerCase();
	}
	public static String UCase(String String1) {
		return String1.toUpperCase();
	}
	public static String Trim(String String1) {
		return String1.trim();
	}
	private static boolean isWhitespace(String val, int index) {
		String whitespaceChars =  " \t\n\r\f\0";
		return (whitespaceChars.indexOf(val.charAt(index)) != -1);
	}
	public static String LTrim(String String1) {
		int iStart = 0;
		for (int i = 0; i < String1.length(); i++) {
			if (isWhitespace(String1, i)) iStart = i + 1; else i = String1.length();
		}
		return String1.substring(iStart);
	}
	public static String RTrim(String String1) {
		int iEnd = String1.length();
		for (int i = String1.length()-1; i >= 0; i--) {
			if (isWhitespace(String1, i)) iEnd = i; else i = 0;
		}
		return String1.substring(0, iEnd);
	}
	public static String String(int Numbers, String TCharacter) {
		byte[] tmpAry = new byte[Numbers]; int iTCharacter = Asc(TCharacter);
		for (int i = 0; i < Numbers; i++) tmpAry[i] = (byte)iTCharacter;
		try {
			return new String(tmpAry, "ISO-8859-1");
		} catch (Exception e) {
			return new String(tmpAry);
		}
	}
	public static String Space(int Numbers) {
		byte[] tmpAry = new byte[Numbers];
		for (int i = 0; i < Numbers; i++) tmpAry[i] = 32;
		try {
			return new String(tmpAry, "ISO-8859-1");
		} catch (Exception e) {
			return new String(tmpAry);
		}
	}
	public static boolean StrCompPattern(String String1, String String2) {
		java.util.regex.Pattern pt = java.util.regex.Pattern.compile(String2, java.util.regex.Pattern.CASE_INSENSITIVE);
		java.util.regex.Matcher mt = pt.matcher(String1);
		return mt.find();
	}
	public static int StrCompText(String String1, String String2) {
		return StrComp(String1, String2, VBCompareMethod.vbTextCompare);
	}
	public static int StrComp(String String1, String String2) {
		return StrComp(String1, String2, VBCompareMethod.vbBinaryCompare);
	}
	public static int StrComp(String String1, String String2, VBCompareMethod CompareMet) {
		int ret = 0;
		if (String1 == null || String2 == null) {
			if (String1 == String2) return 0;
			if (String1 == null) return -1;
			return 1;
		}
		switch ( CompareMet.intValue() ) {
			case ( 0 ): { ret = String1.compareTo(String2); break; }
			case ( 1 ): { ret = String1.compareToIgnoreCase(String2); break; }
			default: { ret = String1.compareToIgnoreCase(String2); break; }
		}
		if (ret > 0) ret = 1;
		return (ret);
	}
	public static String ArrayToString(byte[] strArray) {
		try { return (new String(strArray, "ISO-8859-1")); } catch (Exception e) { try { return (new String(strArray)); } catch (Exception e2) { return ""; } }
	}
	public static byte[] StringToArray(String strVal) {
		byte[] val;
		try { val = strVal.getBytes("ISO-8859-1"); } catch (Exception e) { val = new byte[0]; }
		return (val);
	}
	public static String StrReverse(String String1) {
		byte[] val = StringToArray(String1);
		byte[] val2 = new byte[val.length];
		for (int i = 0; i < val.length; i++) val2[val.length - i - 1] = val[i];
		return (ArrayToString(val2));
	}
        public static String PCase(String String1) {
		String String2 = "";
		boolean lastWasLetter = false;
		for (int i=0; i<=String1.length()-1; i++) {
                        Character ch = Character.valueOf(String1.charAt(i));
                        if (Character.isLetterOrDigit(ch.charValue())) {
                                if (Character.isDigit(ch.charValue())) {
					String2 = String2.concat(ch.toString());
				} else if (lastWasLetter) {
					String2 = String2.concat(ch.toString().toLowerCase());
					lastWasLetter = true;
				} else {
					String2 = String2.concat(ch.toString().toUpperCase());
					lastWasLetter = true;
				}
			} else {
				String2 = String2.concat(ch.toString());
				lastWasLetter = false;
			}
		}
		return (String2); 
	}
	private static String XCase(String String1) {
		return (String1);	
	}
	public static String StrConv(String String1, VBStrConv ConvMet) {
		switch ( ConvMet.intValue() ) {
			case ( 1 ): {	return (UCase(String1));	} // vbUpperCase
			case ( 2 ): {	return (LCase(String1));	} //  vbLowerCase
			case ( 3 ): {	return (PCase(String1));	} //  vbProperCase // NOT IMPLEMENTED
			case ( 4 ): {	return (XCase(String1));	} //  vbWide    // NOT IMPLEMENTED
			case ( 8 ): {	return (XCase(String1));	} //  vbNarrow // NOT IMPLEMENTED
			case ( 16 ): {	return (XCase(String1));	} //  vbKatakana // NOT IMPLEMENTED
			case ( 32 ): {	return (XCase(String1));	} //  vbHiragana // NOT IMPLEMENTED
			case ( 64 ): {	return (XCase(String1));	} //  vbUnicode // NOT IMPLEMENTED
			case ( 128 ): {	return (XCase(String1));	} //  vbFromUnicode // NOT IMPLEMENTED
		}
		return (String1);
	}
	public static VBArrayString Split(String String1) throws Exception {
		VBArrayString ret = new VBArray(); // break to vb
		byte[] myAry = StringToArray(String1);
		ret.setBound(0, myAry.length - 1, false);
		for (int i = 0; i < myAry.length; i++) ret.setValueStr(i, Chr(myAry[i]));
		return (ret);
	}
	public static VBArrayString Split(String String1, String Delimiter) throws Exception {
		return Split(String1, Delimiter, -1);
	}
	public static VBArrayString Split(String String1, String Delimiter, long Limit) throws Exception {
		return Split(String1, Delimiter, Limit, VBCompareMethod.vbBinaryCompare);
	}
	public static VBArrayString Split(String String1, String Delimiter, long Limit, VBCompareMethod CompareMet) throws Exception {
		VBArrayString ret = new VBArray();
		if (Limit == 0) { Limit = -1; } // break to vb but better, in my opinion
		if (Limit == 1) { ret.addStringItem(String1); return ret; }
		int iStart = 1; int iLast = 1; int iMemStart = 0; int iLenDelimiter = Delimiter.length();
		for (int i = 1; ((i < Limit) || (Limit == -1)) && (iStart >= 0); i++) {
			iStart = InStr(iStart, String1, Delimiter, CompareMet);
			if (iStart == 0) {
				iStart = -1;
			} else {
				iMemStart = iStart;
				iStart = iStart + iLenDelimiter;
				ret.addStringItem(Mid(String1, iLast, iStart - iLast - iLenDelimiter));
				iLast = iStart;
			}
		}
		if (iMemStart == 0) { 
			ret.addStringItem(String1); 
		} else {
			ret.addStringItem(Mid(String1, iMemStart + iLenDelimiter, Len(String1) - iMemStart + 1 - iLenDelimiter)); // add rest
		}
		return ret;
	}
	public static String WeekdayName(int Weekday) {
		return (WeekdayName(Weekday, false));
	}	
	public static String WeekdayName(int Weekday, boolean Abbreviate) {
		return (WeekdayName(Weekday, Abbreviate, VBDayOfWeek.vbMonday));
	}
	public static String WeekdayName(int Weekday, boolean Abbreviate, VBDayOfWeek FirstDayOfWeek) {
		if (Weekday < 1 || Weekday > 7) return Constants.vbNullString;
		String[] weekdayName = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"};
		String[] weekdayNameAbbreviate  = {"Mon","Tue","Wed","Thu","Fri","Sat","Sun"};
		int iFirstDayOfWeek = FirstDayOfWeek.intValue() ;
		if (iFirstDayOfWeek == 0) iFirstDayOfWeek = 1; // by system
		int WeekdayIntern = Weekday + iFirstDayOfWeek - 3;
		if (WeekdayIntern < 0) WeekdayIntern = 7 + WeekdayIntern;
		if (WeekdayIntern > 6) WeekdayIntern = WeekdayIntern - 7;
		if (WeekdayIntern >= 0 && WeekdayIntern < 7) {
			if (Abbreviate) return weekdayNameAbbreviate[WeekdayIntern]; else return weekdayName[WeekdayIntern];
		} else {
			return Constants.vbNullString;
		}
	}
	public static String MonthName(int Month) {
		return (MonthName(Month, false));
	}
	public static String MonthName(int Month, boolean Abbreviate) {
		String[] monthName = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
		String[] monthNameAbbreviate = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
		Month = Month - 1;
		if (Month >= 0 && Month < 12) {
			if (Abbreviate) return monthNameAbbreviate[Month]; else return monthName[Month];
		} else {
			return Constants.vbNullString;
		}
	}
	public static String Join(VBArrayString val) throws Exception {
		return (Join(val, Constants.vbNullString));
	}
	public static String Join(VBArrayString val, String Delimiter) throws Exception {
		StringBuffer strBuffer = new StringBuffer();
		for (int i = val.getLBound(); i <= val.getUBound(); i++) {
			strBuffer.append(val.valueOfStr(i));
			if (i < val.getUBound()) if (Delimiter.length() > 0) strBuffer.append(Delimiter);
		}
		return strBuffer.toString();
	}
	public static VBArrayString Filter(VBArrayString val, String Match) throws Exception {
		return Filter(val, Match, true);
	}
	public static VBArrayString Filter(VBArrayString val, String Match, boolean Include) throws Exception {
		return Filter(val, Match, Include, VBCompareMethod.vbBinaryCompare);
	}
	public static VBArrayString Filter(VBArrayString val, String Match, boolean Include, VBCompareMethod CompareMet) throws Exception {
		VBArrayString ret = new VBArray();
		for (int i = val.getLBound(); i <= val.getUBound(); i++) {
			String CurrentItem = val.valueOfStr(i);
			if (InStr(1, CurrentItem, Match, CompareMet) > 0) {
				if (Include) ret.addStringItem(CurrentItem);
			} else {
				if (!Include) ret.addStringItem(CurrentItem);
			}
		}
		return ret;
	}
	
	public static String Format(String Expression, String Format) {
		try {
			java.util.GregorianCalendar tmpCal = VBDateParser.parse(Expression);
			//calendar.get(Calendar.YEAR);
			return tmpCal.toString();
		} catch (Exception e) {
			return e.toString();
		}
	}
	/*
	public static String JSimpleDateFormat(String Expression, String Format) {
		String ret = Constants.vbNullString;
		try {
			java.text.SimpleDateFormat format = new java.text.SimpleDateFormat(Format);
            java.util.Date parsed = format.parse(Expression);
            ret = parsed.toString();
        } catch(Exception pe) {}
		return ret;
	}*/
	
	

/*
		********************* FUNKTIONSLISTE *********************

Asc
AscB
AscW
Chr
Chr$
ChrB
ChrB$
ChrW
ChrW$
Filter
Format
Format$
FormatCurrency
FormatDateTime
FormatNumber
FormatPercent
InStr
InStrB
InStrRev
Join
LCase
LCase$
Left
Left$
LeftB
LeftB$
Len
LenB
LTrim
LTrim$
Mid
Mid$
MidB
MidB$
MonthName
PCase
PCase$
Replace
Right
Right$
RightB
RightB$
RTrim
Space
Space$
Split
StrComp
StrConv
String
String$
StrReverse
Trim
Trim$
UCase
UCase$
WeekdayName
*/
	
	
	
	
	
	
	
	
	
	
	
}
