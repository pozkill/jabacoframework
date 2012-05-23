package VBA;

import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.security.*;
import java.io.*;

/**
 * Cryptography Methods
 * @author Manuel Siekmann
 * @version 1.0
 */
public class Cryptography {

	public static String md5(String s) {
		byte[] defaultBytes = s.getBytes();
		String ret = "";
		try{
			MessageDigest algorithm = MessageDigest.getInstance("MD5");
			algorithm.reset();
			algorithm.update(defaultBytes);
			byte messageDigest[] = algorithm.digest();
			StringBuffer hexString = new StringBuffer();
			for (int i=0; i<messageDigest.length; i++) {
				if(messageDigest[i] <= 15 && messageDigest[i] >= 0) {
					hexString.append("0");
				}
				hexString.append(Integer.toHexString(0xFF & messageDigest[i]));
			}
			ret = hexString.toString();
		} catch (Exception e) {}
		if (ret == null) ret = "";
		return ret;
	}

	public static String sha1(String s) {
		byte[] defaultBytes = s.getBytes();
		String ret = "";
		try{
			MessageDigest algorithm = MessageDigest.getInstance("SHA-1");
			algorithm.reset();
			algorithm.update(defaultBytes);
			byte messageDigest[] = algorithm.digest();
			StringBuffer hexString = new StringBuffer();
			for (int i=0; i<messageDigest.length; i++) {
				if(messageDigest[i] <= 15 && messageDigest[i] >= 0) {
					hexString.append("0");
				}
				hexString.append(Integer.toHexString(0xFF & messageDigest[i]));
			}
			ret = hexString.toString();
		} catch (Exception e) {}
		if (ret == null) ret = "";
		return ret;
	}

	private static char[] base64map1 = new char[64]; static { int i = 0; for (char c='A'; c<='Z'; c++) base64map1[i++] = c; for (char c='a'; c<='z'; c++) base64map1[i++] = c; for (char c='0'; c<='9'; c++) base64map1[i++] = c; base64map1[i++] = '+'; base64map1[i++] = '/'; }
	private static byte[] base64map2 = new byte[128]; static { for (int i = 0; i<base64map2.length; i++) base64map2[i] = -1; for (int i=0; i<64; i++) base64map2[base64map1[i]] = (byte)i; }

	public static String encodeBase64 (String s) {
		byte[] in = s.getBytes();
		int iLen = in.length;
		int oDataLen = (iLen*4+2)/3;       // output length without padding
		int oLen = ((iLen+2)/3)*4;         // output length including padding
		char[] out = new char[oLen];
		int ip = 0;
		int op = 0;
		while (ip < iLen) {
		  int i0 = in[ip++] & 0xff;
		  int i1 = ip < iLen ? in[ip++] & 0xff : 0;
		  int i2 = ip < iLen ? in[ip++] & 0xff : 0;
		  int o0 = i0 >>> 2;
		  int o1 = ((i0 &   3) << 4) | (i1 >>> 4);
		  int o2 = ((i1 & 0xf) << 2) | (i2 >>> 6);
		  int o3 = i2 & 0x3F;
		  out[op++] = base64map1[o0];
		  out[op++] = base64map1[o1];
		  out[op] = op < oDataLen ? base64map1[o2] : '='; op++;
		  out[op] = op < oDataLen ? base64map1[o3] : '='; op++; 
		}
		return new String(out); 
	}

	public static String decodeBase64 (String s) {
		char[] in = s.toCharArray();
		int iLen = in.length;
		if (iLen%4 != 0) throw new IllegalArgumentException ("Length of Base64 encoded input string is not a multiple of 4.");
		while (iLen > 0 && in[iLen-1] == '=') iLen--;
		int oLen = (iLen*3) / 4;
		byte[] out = new byte[oLen];
		int ip = 0;
		int op = 0;
		while (ip < iLen) {
			int i0 = in[ip++];
			int i1 = in[ip++];
			int i2 = ip < iLen ? in[ip++] : 'A';
			int i3 = ip < iLen ? in[ip++] : 'A';
			if (i0 > 127 || i1 > 127 || i2 > 127 || i3 > 127)
				 throw new IllegalArgumentException ("Illegal character in Base64 encoded data.");
			int b0 = base64map2[i0];
			int b1 = base64map2[i1];
			int b2 = base64map2[i2];
			int b3 = base64map2[i3];
			if (b0 < 0 || b1 < 0 || b2 < 0 || b3 < 0)
				 throw new IllegalArgumentException ("Illegal character in Base64 encoded data.");
			int o0 = ( b0       <<2) | (b1>>>4);
			int o1 = ((b1 & 0xf)<<4) | (b2>>>2);
			int o2 = ((b2 &   3)<<6) |  b3;
			out[op++] = (byte)o0;
			if (op<oLen) out[op++] = (byte)o1;
			if (op<oLen) out[op++] = (byte)o2; 
		}
		return new String(out);
	}
}
