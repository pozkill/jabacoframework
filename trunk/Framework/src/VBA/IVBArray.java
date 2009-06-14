package VBA;

public interface IVBArray {

	public byte[] byteArray() throws Exception;
	public boolean[] boolArray() throws Exception;
	public int[] intArray() throws Exception;
	public double[] doubleArray() throws Exception;
	public long[] longArray() throws Exception;
	public float[] floatArray() throws Exception;
	public String[] stringArray() throws Exception;
	public Object[] objectArray() throws Exception;	
	
	public int getLBound() throws Exception;
	public int getLBound(int Index) throws Exception; 
	public int getUBound() throws Exception;
	public int getUBound(int Index) throws Exception; 
	
	public int getBase();
	public void setBound(int iNewLBound, int iNewUBound, boolean bPreserve) throws Exception;
	public VBArray getFromDimension(int Index) throws Exception;
	public void addDimension(int iNewLBound, int iNewUBound, boolean bPreserve) throws Exception;
	
	public Object clone();
	public VBArray cloneArray() throws Exception;
	public Object valueOf(int Index) throws Exception;
	
	public void setValueNative(int Index, Object Value) throws Exception;
	public void setValue(int Index, Object Value) throws Exception;

	public double valueOfDbl(int Index) throws Exception;
	public void setValueDbl(int Index, double Value) throws Exception;

	public int valueOfInt(int Index) throws Exception;
	public void setValueInt(int Index, int Value) throws Exception;

	public long valueOfLng(int Index) throws Exception;
	public void setValueLng(int Index, long Value) throws Exception;
	
	public Object valueOfObj(int Index) throws Exception;
	public void setValueObj(int Index, Object Value) throws Exception;
	
	public void setValueVar(int Index, VBVariant Value) throws Exception;
	public VBVariant valueOfVar(int Index) throws Exception;
	
	public void addStringItem(String NewItem) throws Exception;

	public void setValueStr(int Index, String Value) throws Exception;
	public String valueOfStr(int Index) throws Exception;
	
	public String toDebugString() throws Exception;	
	public String toString();
	
	
}

