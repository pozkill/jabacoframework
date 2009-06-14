package VBA;
public class VBCallType extends VBEnumClass {
	public static VBCallType VBMethod  = new VBCallType(1);
	public static VBCallType VBGet  = new VBCallType(2);
	public static VBCallType VBLet  = new VBCallType(4);
	public static VBCallType VBSet = new VBCallType(8);

	public VBCallType () {}
	public VBCallType (int iVBCallType) {
		super(iVBCallType);
	}
}
