package VBA;

public class VBVarType extends VBEnumClass {
	public static VBVarType vbEmpty = new VBVarType(0);
	public static VBVarType vbNull = new VBVarType(1);
	public static VBVarType vbInteger = new VBVarType(2);
	public static VBVarType vbLong = new VBVarType(3);
	public static VBVarType vbSingle = new VBVarType(4);
	public static VBVarType vbDouble = new VBVarType(5);
	public static VBVarType vbCurrency = new VBVarType(6);
	public static VBVarType vbDate = new VBVarType(7);
	public static VBVarType vbString = new VBVarType(8);
	public static VBVarType vbObject = new VBVarType(9);
	public static VBVarType vbError = new VBVarType(10);
	public static VBVarType vbBoolean = new VBVarType(11);
	public static VBVarType vbVariant = new VBVarType(12);
	public static VBVarType vbDataObject = new VBVarType(13);
	public static VBVarType vbDecimal = new VBVarType(14);
	public static VBVarType vbByte = new VBVarType(17);
	public static VBVarType vbUserDefinedType = new VBVarType(36);
	public static VBVarType vbArray = new VBVarType(8192);
	public VBVarType () {}
	public VBVarType (int iVBVarType) {
		super(iVBVarType);
	}

}
