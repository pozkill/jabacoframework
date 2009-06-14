package VBA;

public class VBAppWinStyle extends VBEnumClass {
	public static VBAppWinStyle vbHide = new VBAppWinStyle(0);
	public static VBAppWinStyle vbNormalFocus = new VBAppWinStyle(1);
	public static VBAppWinStyle vbMinimizedFocus = new VBAppWinStyle(2);
	public static VBAppWinStyle vbMaximizedFocus = new VBAppWinStyle(3);
	public static VBAppWinStyle vbNormalNoFocus = new VBAppWinStyle(4);
	public static VBAppWinStyle vbMinimizedNoFocus = new VBAppWinStyle(6);

	public VBAppWinStyle () {}
	public VBAppWinStyle (int iVBAppWinStyle) {
		super(iVBAppWinStyle);
	}
}
