package VBA;

public class VBMsgBoxStyle extends VBEnumClass {
	public static VBMsgBoxStyle vbApplicationModal = new VBMsgBoxStyle(0);
	public static VBMsgBoxStyle vbAbortRetryIgnore = new VBMsgBoxStyle(2);
	public static VBMsgBoxStyle vbCritical = new VBMsgBoxStyle(16);
	public static VBMsgBoxStyle vbDefaultButton1 = new VBMsgBoxStyle(0);
	public static VBMsgBoxStyle vbDefaultButton2 = new VBMsgBoxStyle(256);
	public static VBMsgBoxStyle vbDefaultButton3 = new VBMsgBoxStyle(512);
	public static VBMsgBoxStyle vbDefaultButton4 = new VBMsgBoxStyle(768);
	public static VBMsgBoxStyle vbExclamation = new VBMsgBoxStyle(48);
	public static VBMsgBoxStyle vbInformation = new VBMsgBoxStyle(64);
	public static VBMsgBoxStyle vbMsgBoxHelpButton = new VBMsgBoxStyle(16384);
	public static VBMsgBoxStyle vbMsgBoxRight = new VBMsgBoxStyle(524288);
	public static VBMsgBoxStyle vbMsgBoxRtlReading = new VBMsgBoxStyle(1048576);
	public static VBMsgBoxStyle vbMsgBoxSetForeground = new VBMsgBoxStyle(65536);
	public static VBMsgBoxStyle vbOKCancel = new VBMsgBoxStyle(1);
	public static VBMsgBoxStyle vbOKOnly = new VBMsgBoxStyle(0);
	public static VBMsgBoxStyle vbQuestion = new VBMsgBoxStyle(32);
	public static VBMsgBoxStyle vbRetryCancel = new VBMsgBoxStyle(5);
	public static VBMsgBoxStyle vbSystemModal = new VBMsgBoxStyle(4096);
	public static VBMsgBoxStyle vbYesNo = new VBMsgBoxStyle(4);
	public static VBMsgBoxStyle vbYesNoCancel = new VBMsgBoxStyle(3);
	public VBMsgBoxStyle () {}
	public VBMsgBoxStyle (int iVBMsgBoxStyle) {
		super(iVBMsgBoxStyle);
	}
}
