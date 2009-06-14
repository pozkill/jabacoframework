package VBA;

public class VBCompareMethod extends VBEnumClass {
	//public static final int VB_BINARY_COMPARE = 0;
	//public static final int VB_TEXT_COMPARE = 1;

	public static VBCompareMethod vbBinaryCompare = new VBCompareMethod (0);
	public static VBCompareMethod vbTextCompare = new VBCompareMethod(1);
	//public static VBCompareMethod vbDatabaseCompare = new VBCompareMethod (2);

	public VBCompareMethod () {}
	public VBCompareMethod (int iVBCompareMethod) {
		super(iVBCompareMethod);
	}
}
