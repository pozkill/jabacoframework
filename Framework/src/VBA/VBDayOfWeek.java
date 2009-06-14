package VBA;

public class VBDayOfWeek extends VBEnumClass {
	public static VBDayOfWeek vbUseSystemDayOfWeek  = new VBDayOfWeek(0);
	public static VBDayOfWeek vbSunday  = new VBDayOfWeek(1);
	public static VBDayOfWeek vbMonday  = new VBDayOfWeek(2);
	public static VBDayOfWeek vbTuesday  = new VBDayOfWeek(3);
	public static VBDayOfWeek vbWednesday  = new VBDayOfWeek(4);
	public static VBDayOfWeek vbThursday  = new VBDayOfWeek(5);
	public static VBDayOfWeek vbFriday  = new VBDayOfWeek(6);
	public static VBDayOfWeek vbSaturday  = new VBDayOfWeek(7);
	public VBDayOfWeek () {}
	public VBDayOfWeek (int iVBDayOfWeek) {
		super(iVBDayOfWeek);
	}
}
