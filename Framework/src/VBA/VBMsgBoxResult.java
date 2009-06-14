package VBA;

public class VBMsgBoxResult extends VBEnumClass {
	public static VBMsgBoxResult vbOK = new VBMsgBoxResult(1);
	public static VBMsgBoxResult vbCancel = new VBMsgBoxResult(2);
	public static VBMsgBoxResult vbAbort = new VBMsgBoxResult(3);
	public static VBMsgBoxResult vbRetry = new VBMsgBoxResult(4);
	public static VBMsgBoxResult vbIgnore = new VBMsgBoxResult(5);
	public static VBMsgBoxResult vbYes = new VBMsgBoxResult(6);
	public static VBMsgBoxResult vbNo = new VBMsgBoxResult(7);
	public VBMsgBoxResult () {}
	public VBMsgBoxResult (int iVBMsgBoxResult) {
		super(iVBMsgBoxResult);
	}
}
