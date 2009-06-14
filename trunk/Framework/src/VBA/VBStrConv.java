package VBA;

public class VBStrConv extends VBEnumClass {
	public static VBStrConv vbUpperCase = new VBStrConv(1);
	public static VBStrConv vbLowerCase = new VBStrConv(2);
	public static VBStrConv vbProperCase = new VBStrConv(3);
	public static VBStrConv vbWide = new VBStrConv(4);
	public static VBStrConv vbNarrow = new VBStrConv(8);
	public static VBStrConv vbKatakana = new VBStrConv(16);
	public static VBStrConv vbHiragana = new VBStrConv(32);
	public static VBStrConv vbUnicode = new VBStrConv(64);
	public static VBStrConv vbFromUnicode = new VBStrConv(128);
	public VBStrConv () {}
	public VBStrConv (int iVBStrConv) {
		super(iVBStrConv);
	}
}
