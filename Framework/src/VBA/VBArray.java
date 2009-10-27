package VBA;
import java.io.Serializable;
import java.util.*;

public class VBArray implements Cloneable, Serializable, VBArrayString, VBArrayVariant, VBArrayChar, VBArrayByte, VBArrayDate, VBArrayDouble, VBArrayInteger, VBArrayLong, VBArrayBoolean, VBArraySingle, VBArrayObject, IVBArray {
	Class myClass;
	int iUBound = -1;
	int iLBound = 0;
	int iDimension = 0;
	Object[] colObject;
	
	/* tmpCollections ... */
	boolean writeBack = false;
	byte[] colByte;
	int[] colInteger;
	double[] colDouble;
	long[] colLong;
	float[] colFloat;
	boolean[] colBoolean;
	
	/* ... tmpCollections */

	int LastIndex = -1;
	public VBArray() {}

	
	/********** CREATE JAVA-ARRAY ***********/
	public byte[] byteArray() throws Exception { // create bytestream 
		writeBackArray();
		if (colObject == null) return null;
		colByte = new byte[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colByte[i] = (Conversion.CByte(colObject[i]));
		}
		writeBack = true;
		return colByte;
	}
	public boolean[] boolArray() throws Exception { // create bytestream 
		writeBackArray();
		if (colObject == null) return null;
		colBoolean = new boolean[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colBoolean[i] = (boolean)((Conversion.CInt(colObject[i])) != 0);
		}
		writeBack = true;
		return colBoolean;
	}	
	public int[] intArray() throws Exception { // create intstream !
		writeBackArray();
		if (colObject == null) return null;
		colInteger = new int[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colInteger[i] = (Conversion.CInt(colObject[i]));
		}
		writeBack = true;
		return colInteger;
	}
	public double[] doubleArray() throws Exception { // create doublestream 
		writeBackArray();
		if (colObject == null) return null;
		colDouble = new double[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colDouble[i] = Conversion.CDbl(colObject[i]);
		}
		writeBack = true;
		return colDouble;
	}
	public long[] longArray() throws Exception { // create longstream 
		writeBackArray();
		if (colObject == null) return null;
		colLong = new long[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colLong[i] = Conversion.CLng(colObject[i]);
		}
		writeBack = true;
		return colLong;
	}
	public float[] floatArray() throws Exception { // create floatstream !
		writeBackArray();
		if (colObject == null) return null;
		colFloat = new float[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colFloat[i] = (new Float((Conversion.CDbl(colObject[i])))).floatValue();
		}
		writeBack = true;
		return colFloat;
	}
	public String[] stringArray() throws Exception { // return objectarray
		String[] colString;
		if (colObject == null) return null;
		colString = new String[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			colString[i] = new String((String)(colObject[i]));
		}
		return colString;
	}
	
	public Object[] objectArray() throws Exception { // return objectarray
		return colObject;
	}	
	
	public Object[] objectArrayClone() {
		Object[] colClone;
		if (colObject == null) return null;
		colClone = new Object[colObject.length];
		for (int i = 0; i < colObject.length; i++) {
			if (colObject[i] instanceof Boolean) {			colClone[i] = new Boolean(((Boolean)colObject[i]).booleanValue());
			}else if(colObject[i] instanceof Byte) {		colClone[i] = new Byte(((Byte)colObject[i]).byteValue());
			}else if(colObject[i] instanceof Long) {		colClone[i] = new Long(((Long)colObject[i]).longValue());
			}else if(colObject[i] instanceof Integer) {		colClone[i] = new Integer(((Integer)colObject[i]).intValue());
			}else if(colObject[i] instanceof Double) {		colClone[i] = new Double(((Double)colObject[i]).doubleValue());
			}else if(colObject[i] instanceof Float) {		colClone[i] = new Float(((Float)colObject[i]).floatValue());
			}else if(colObject[i] instanceof String) {		colClone[i] = new String(((String)colObject[i]).toString());
			} else {										
				colClone[i] = colObject[i];
			}
		}
		return colClone;
	}
	
	public Object clone() {
		try {
			return (super.clone());
		} catch (Exception e) {
			return this;
		}
	}
	
	public VBArray cloneSimple() throws Exception {
		Object[] obj = this.objectArrayClone();
		VBArray ret = new VBArray();
		if (obj == null) {
			ret = new VBArray();
		} else {
			ret = new VBArray(obj);
		}
		return ret;
	}
	
	public VBArray cloneArray() throws Exception {
		if (iDimension == 0) {
			return cloneSimple();
		} else {
			VBArray ret = new VBArray(iLBound, iUBound, iDimension);
			for (int i = 0; i < colObject.length; i++) {
				ret.setValueNative(i, ((VBArray)(colObject[i])).cloneArray());
			}
			return ret;
		}
	}

	private void writeBackArray() {  // writeback bytestream 
		if (writeBack == false) return;
		for (int i = 0; i < colObject.length; i++)	{
			if (colByte != null) { if (colByte.length == colObject.length) {
				colObject[i] = new Integer(colByte[i]);
			}}
			if (colInteger != null) { if (colInteger.length == colObject.length) {
				colObject[i] = new Integer(colInteger[i]);
			}}
			if (colDouble != null) { if (colDouble.length == colObject.length) {
				colObject[i] = new Double(colDouble[i]);
			}}
			if (colLong != null) { if (colLong.length == colObject.length) {
				colObject[i] = new Long(colLong[i]);
			}}
			if (colFloat != null) { if (colFloat.length == colObject.length) {
				colObject[i] = new Double(colFloat[i]);
			}}
			if (colBoolean != null) { if (colBoolean.length == colObject.length) {
				colObject[i] = new Integer(colBoolean[i] ? 1 : 0);
			}}
		}
		clearWriteBackArray();
	}
	
	private void clearWriteBackArray() {
		colByte = null;
		colInteger = null;
		colDouble = null;
		colLong = null;
		colFloat = null;
		colBoolean = null;
		writeBack = false;
	}

    public Iterator iterator() {
		return new VBArrayIterator(colObject);
    }
	
	/********** VBARRAY METHODS ***********/
	
	public Object valueOf(int Index) throws Exception {
		//try {
			writeBackArray();
			LastIndex = calcOffset(Index);
			//if (Index >= iLBound && Index <= iLBound) {
			if (myClass != null) {
				try {
					if (colObject[LastIndex] == null) colObject[LastIndex] = myClass.newInstance();
				} catch (Exception ex) {}
			}	

			return colObject[LastIndex];
			//} else {
			//	throw new java.lang.ArrayIndexOutOfBoundsException("Array Index Out Of Bounds!");
			//}
		//} catch (Exception e) {
		//	return null;
		//}
	}
	
	public void setValueNative(int Index, Object Value) {
		colObject[Index] = Value;
	}
	
	public void setValue(int Index, Object Value) throws Exception {
		writeBackArray();
		try {
			LastIndex = calcOffset(Index);
			colObject[calcOffset(Index)] = Value;
		} catch (Exception e) {
			throw new java.lang.ArrayIndexOutOfBoundsException();
		}		
	}
	
	public int getLBound() throws Exception { return getLBound(1); }
	public int getLBound(int Index) throws Exception { 
		return getLBoundInternal((iDimension + 1) - Index);
	}
	private int getLBoundInternal (int Index) throws Exception {
		try {
			if (Index < 0 || Index > iDimension) return -1;
			if (Index == 0) return iLBound; 
			return ((VBArray)colObject[0]).getLBoundInternal(Index-1);
		} catch (Exception e) { }
		return -1;
	}
	public int getUBound() throws Exception { return getUBound(1); }
	public int getUBound(int Index)  throws Exception { 
		return getUBoundInternal((iDimension + 1) - Index);
	}
	public int getUBoundInternal(int Index)  throws Exception {
		try {
			if (Index < 0 || Index > iDimension) return -1;
			if (Index == 0) return iUBound; // reverse
			return ((VBArray)colObject[0]).getUBoundInternal(Index-1);
		} catch (Exception e) { }
		return -1;
	}

	/* deprecated */
	public int getBase() { 
		return 0;			
	}
	
	public void setBound(int iNewLBound, int iNewUBound, boolean bPreserve) throws Exception {
		writeBackArray();
		iLBound = iNewLBound;
		iUBound = iNewUBound;
		Object[] tmpArray = new Object[iUBound + -iLBound + 1];
		if (bPreserve) {
			if (colObject != null) {
				int iLen = colObject.length;
				if (iLen > tmpArray.length) iLen = tmpArray.length;
				System.arraycopy (colObject,0,tmpArray,0,iLen);
			}
		}
		colObject = tmpArray;
	}
	
	public VBArray getFromDimension(int Index) throws Exception {
		if (iDimension != 0) {
			return (VBArray)valueOf(Index);
		} else {
			return null;
		}
	}
	
	public void addDimension(int iNewLBound, int iNewUBound, boolean bPreserve) throws Exception {
		writeBackArray();
		if (iDimension == 0) {
			for (int i = 0; i < colObject.length; i++) {
				colObject[i] = new VBArray(myClass);
				((VBArray)colObject[i]).setBound(iNewLBound, iNewUBound, bPreserve);
			}
		} else {
			for (int i = 0; i < colObject.length; i++) {
				VBArray tmp = (VBArray)colObject[i];
				tmp.addDimension (iNewLBound, iNewUBound, bPreserve);
			}
		}
		iDimension++;
	}
	
	private int calcOffset(int Index) {
		return Index - iLBound;
	}
	private int calcOffsetReverse(int Index) {
		return Index + iLBound;
	}
	
	public double valueOfDbl(int Index) throws Exception {
		return Conversion.CDbl(valueOf(Index));
	}
	
	public void setValueDbl(int Index, double Value) throws Exception {
		setValue (Index, new Double(Value));
	}
	
	public int valueOfInt(int Index) throws Exception {
		return Conversion.CInt(valueOf(Index));
	}
	
	public void setValueInt(int Index, int Value) throws Exception {
		setValue (Index, new Integer(Value));
	}
	
	public long valueOfLng(int Index) throws Exception {
		return Conversion.CLng(valueOf(Index));
	}
	
	public void setValueLng(int Index, long Value) throws Exception {
		setValue (Index, new Long(Value));
	}
	
	public Object valueOfObj(int Index) throws Exception {
		try {
			return valueOf(Index);
		} catch (Exception e) { return null; }
	}
	
	public void setValueObj(int Index, Object Value) throws Exception {
		setValue (Index, Value);
	}
	
	public String valueOfStr(int Index) throws Exception {
		return Conversion.CStr(valueOf(Index));
	}
	
	public VBVariant valueOfVar(int Index) throws Exception {
		try {
			if (valueOf(Index) instanceof VBVariant) {
				return (VBVariant)(valueOf(Index));
			} else {
				return new VBVariant(valueOf(Index));
			}
		} catch (Exception e) { 
			setValueVar(Index, new VBVariant());
			try { // retry
				return (VBVariant)valueOf(Index);
			} catch (Exception e2) { 
				return null;
			}
		}
	}
	
	public void setValueVar(int Index, VBVariant Value) throws Exception {
		setValue (Index, Value);
	}
	
	public void setValueStr(int Index, String Value) throws Exception {
		setValue (Index, Value);
	}
	
	public void addStringItem(String NewItem) throws Exception {
		int iNewUBound = this.getUBound() + 1;
		setBound(this.getLBound(), iNewUBound, true);
		setValueStr(iNewUBound, NewItem);
	}
	
	/********** TO STRING ***********/
	
	public String toDebugString() {
		writeBackArray();
		try {
			if ((iUBound - iLBound) < 0) return "Empty";
			if (colObject == null) return "Null";
			String ret = "";
			String tmp = "";
			int myIndex = LastIndex - 3;
			if (myIndex < 0) myIndex = 0;
			int myIndexMax = myIndex + 7;
			if (myIndexMax > colObject.length) myIndexMax = colObject.length;
			for (int i = myIndex; i < myIndexMax; i++) {
				if (colObject[i] != null) {
					tmp = colObject[i].toString().replaceAll("\n", "\\n");
					if (tmp.length() > 25) {
						tmp = tmp.substring(0, 25);
					}					
					ret = ret + "|VAR|(" + (new Integer(calcOffsetReverse(i))).toString() + "): " + tmp + "\n";
				} else {
					ret = ret + "|VAR|(" + (new Integer(calcOffsetReverse(i))).toString() + "): empty\n";
				}
			}
			return ret;
		} catch (Exception e) {
			return ("DBGERR: " + e.toString());
		}
	}
	
	
	
	
	public String toString() {
		String ret = "";
		try {
			ret = (new String(byteArray(), "ISO-8859-1"));
		} catch (Exception e) {}
		clearWriteBackArray(); // the array is only for the string 
		return ret;
	}

	
	
	/********** VBARRAY CONSTRUCTORS ***********/
	public VBArray(Class refClass) {
		try {
			myClass = refClass;
		} catch (Exception e) {}
	}
	public VBArray(int refLBound, int refUBound, int refDimension) throws Exception {
		setBound(refLBound, refUBound, false);
		iDimension = refDimension;
	}
	public VBArray(String strVal) throws Exception {
		byte[] val = Strings.StringToArray(strVal);
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueInt(i, val[i]);
	}
	public VBArray(String[] val) throws Exception {
		setBound(0, val.length-1, false);
		System.arraycopy (val,0,colObject,0,val.length);
	}
	public VBArray(long[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueLng(i, val[i]);
	}
	public VBArray(int[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueInt(i, val[i]);
	}
	public VBArray(double[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueDbl(i, val[i]);
	}
	public VBArray(float[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueDbl(i, val[i]);
	}
	public VBArray(byte[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueInt(i, val[i]);
	}
	public VBArray(Object[] val) throws Exception {
		setBound(0, val.length-1, false);
		System.arraycopy (val,0,colObject,0,val.length);
	}
	public VBArray(boolean[] val) throws Exception {
		setBound(0, val.length-1, false);
		for (int i = 0; i <= val.length-1; i++)	setValueInt(i, (val[i] == false) ? 0 : -1);
	}
	public VBArray(java.util.Date[] val) throws Exception {
		setBound(0, val.length-1, false);
		System.arraycopy (val,0,colObject,0,val.length);
	}
	public VBArray(VBVariant[] val) throws Exception {
		setBound(0, val.length-1, false);
		System.arraycopy (val,0,colObject,0,val.length);
	}

	
	/********************* STATIC METHODS *********************/
	/*public static long LBound(VBArray val) {
		return (val.getLBound());
	}	
	public static long UBound(VBArray val) {
		return (val.getUBound());
	}*/
	
	/********************* STATIC CREATE VBARRY *********************/
	public static VBArrayVariant createParamArray(int LBound, int UBound) throws Exception {
		VBArrayVariant ret = new VBArray();
		if (UBound > -1) ret.setBound(LBound, UBound, false);
		return ret;
	}
	public static VBArrayVariant createParamArray(int UBound) throws Exception {
		VBArrayVariant ret = new VBArray();
		if (UBound > -1) ret.setBound(0, UBound, false);
		return ret;
	}
	public static VBArray createVBArray(Class val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayByte createVBArray(String strVal) throws Exception {
		return (new VBArray(Strings.StringToArray(strVal)));
	}	
	public static VBArray createVBArray(Object[] objVal) throws Exception {
		return (new VBArray(objVal));
	}	
	public static VBArrayString createVBArray(String[] strVal) throws Exception {
		return (new VBArray(strVal));
	}	
	public static VBArrayInteger createVBArray(int[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayBoolean createVBArray(boolean[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArraySingle createVBArray(float[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayLong createVBArray(long[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayDouble createVBArray(double[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayByte createVBArray(byte[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayDate createVBArray(java.util.Date[] val) throws Exception {
		return (new VBArray(val));
	}	
	public static VBArrayVariant createVBArray(VBVariant[] val) throws Exception {
		return (new VBArray(val));
	}	
	
	
	
	
}

