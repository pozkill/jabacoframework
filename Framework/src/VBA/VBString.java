package VBA;

import java.lang.StringBuffer;

public class VBString {

	public static String appendStrings(String s1, String s2) {
		StringBuffer myStringBuffer = new StringBuffer();
		myStringBuffer.append (s1);
		myStringBuffer.append (s2);
		return myStringBuffer.toString();
	}

	public static String addString(String s1, String s2) {
		double d1 = 0; double d2 = 0;
		try {
			d1 = Double.valueOf(s1).doubleValue();
			d2 = Double.valueOf(s2).doubleValue();
		} catch (Exception e) {
			return (appendStrings(s1, s2));
		}
		return (new Double(d1 + d2).toString());
	}
	/*public static String str2Double(String s1) {
		String s2 = "";
		for (int i = 0; i <= s1.length(); i++) {
			if (int)s1.charAt(i) >= (int)"0"
		}
	}*/
	public static String divString(String s1, String s2) {
		double d1 = 0;
		double d2 = 0;
		try {
			d1 = Double.valueOf(s1).doubleValue();
		} catch (Exception e) {}
		try {
			d2 = Double.valueOf(s2).doubleValue();
		} catch (Exception e) {}
		return (new Double(d1 / d2).toString());
	}
	public static String mulString(String s1, String s2) {
		double d1 = 0;
		double d2 = 0;
		try {
			d1 = Double.valueOf(s1).doubleValue();
		} catch (Exception e) {}
		try {
			d2 = Double.valueOf(s2).doubleValue();
		} catch (Exception e) {}
		return (new Double(d1 * d2).toString());
	}
	public static String subString(String s1, String s2) {
		double d1 = 0;
		double d2 = 0;
		try {
			d1 = Double.valueOf(s1).doubleValue();
		} catch (Exception e) {}
		try {
			d2 = Double.valueOf(s2).doubleValue();
		} catch (Exception e) {}
		return (new Double(d1 - d2).toString());
	}
	public static String modString(String s1, String s2) {
		double d1 = 0;
		double d2 = 0;
		try {
			d1 = Double.valueOf(s1).doubleValue();
		} catch (Exception e) {}
		try {
			d2 = Double.valueOf(s2).doubleValue();
		} catch (Exception e) {}
		return (new Double(d1 % d2).toString());
	}
	public static String andString(String s1, String s2) {
		long l1 = 0;
		long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		try {
			l2 = Long.valueOf(s2).longValue();
		} catch (Exception e) {}
		return (new Long(l1 & l2).toString());
	}
	public static String orString(String s1, String s2) {
		long l1 = 0;
		long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		try {
			l2 = Long.valueOf(s2).longValue();
		} catch (Exception e) {}
		return (new Long(l1 | l2).toString());
	}
	public static String xorString(String s1, String s2) {
		long l1 = 0;
		long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		try {
			l2 = Long.valueOf(s2).longValue();
		} catch (Exception e) {}
		return (new Long(l1 ^ l2).toString());
	}
	public static String notString(String s1, String s2) {
		long l1 = 0;
		//long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		//try {
		//	l2 = Long.valueOf(s2).longValue();
		//} catch (Exception e) {}
		return (new Long((l1 + 1) * -1).toString());
	}
	public static String shiftlString(String s1, String s2) {
		long l1 = 0;
		long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		try {
			l2 = Long.valueOf(s2).longValue();
		} catch (Exception e) {}
		return (new Long(l1 << l2).toString());
	}
	public static String shiftrString(String s1, String s2) {
		long l1 = 0;
		long l2 = 0;
		try {
			l1 = Long.valueOf(s1).longValue();
		} catch (Exception e) {}
		try {
			l2 = Long.valueOf(s2).longValue();
		} catch (Exception e) {}
		return (new Long(l1 >> l2).toString());
	}
	public static String getString(String s1) {
		if (s1 == null) return "";
		return s1;
	}
	public static boolean equalStringsGL(String s1, String s2) {
		if (s2 == null || s1 == null) {
			if (s1 == s2) return true; else return false;
		}
		s1 = getString(s1); s2 = getString(s2);
		return (s1.equals(s2));
	}
	public static boolean equalStringsUG(String s1, String s2) {
		if (s2 == null || s1 == null) {
			if (s1 != s2) return true; else return false;
		}
		s1 = getString(s1); s2 = getString(s2);
		return (!(s1.equals(s2)));
	}
	public static boolean equalStringsGG(String s1, String s2) {
		s1 = getString(s1); s2 = getString(s2);
		return (s1.compareTo(s2) >= 0);
	}
	public static boolean equalStringsGR(String s1, String s2) {
		s1 = getString(s1); s2 = getString(s2);
		return (s1.compareTo(s2) > 0);
	}
	public static boolean equalStringsKL(String s1, String s2) {
		s1 = getString(s1); s2 = getString(s2);
		return (s1.compareTo(s2) < 0);
	}
	public static boolean equalStringsKG(String s1, String s2) {
		s1 = getString(s1); s2 = getString(s2);
		return (s1.compareTo(s2) <= 0);
	}

}
