package VBA;
import java.util.*;
import java.text.SimpleDateFormat;
import java.text.DateFormat;

	/*
		Function Weekday(Date, [FirstDayOfWeek As VbDayOfWeek = vbSunday])
		Function TimeValue(Time As String)
		Function TimeSerial(Hour As Integer, Minute As Integer, Second As Integer)
		Calendar
		Date
		Date$
		DateAdd
		DateDiff
		DatePart
		DateSerial
		DateValue
		Time
		Time$
		TimeSerial
	*/

public class DateTime {
	/*
	Returns an Integer value from 1 through 9999 representing the year.
	@author Manuel Siekmann
	@param  Date value from which you want to extract the year.
	*/
	public static VBVariant Year(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.YEAR));
	}
	/*
	Returns an Integer value from 1 through 12 representing the month of the year.
	@author Manuel Siekmann
	@param  Date value from which you want to extract the month.
	*/
	public static VBVariant Month(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.MONTH) + 1);
		
		//2009_08_03 OlimilO: why not simply Date.getMonth + 1 ???
		
	}
	/*
	Returns an Integer value from 1 through 31 representing the day of the month.
	@author Manuel Siekmann
	@param  Date value from which you want to extract the month.
	*/
	public static VBVariant Day(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.DAY_OF_MONTH));
	}
	/*
	Returns an Integer value from 0 through 23 that represents the hour of the day. 
	@author Manuel Siekmann
	@param  Date value from which you want to extract the hour.
	*/
	public static VBVariant Hour(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.HOUR_OF_DAY));
	}
	/*
	Returns an Integer value from 0 through 59 that represents the minute of the hour. 
	@author Manuel Siekmann
	@param  Date value from which you want to extract the minute.
	*/
	public static VBVariant Minute(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.MINUTE));
	}
	/*
	Returns an Integer value from 0 through 59 that represents the second of the minute.
	@author Manuel Siekmann
	@param  Date value from which you want to extract the second.
	*/
	public static VBVariant Second(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.SECOND));
	}
	/*
	Returns an Integer value containing a number representing the day of the week.
	@author Manuel Siekmann
	@param  Date value for which you want to determine the day of the week.
	*/
	public static VBVariant Weekday(Date DateValue) {
		Calendar cal = Calendar.getInstance(); cal.setTime(DateValue);
		return new VBVariant(cal.get(Calendar.DAY_OF_WEEK));
	}
	/*
	Returns a Long value representing the current time in milliseconds. 
	@author Manuel Siekmann
	*/
	public static long Timer() {
		return (java.lang.System.currentTimeMillis());
	}
	/*
	Returns a Date value containing the current date and time. Format: yyyy/MM/dd HH:mm:ss
	@author Manuel Siekmann
	*/
	public static VBVariant Now() {
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Calendar cal = Calendar.getInstance(); 
        Date dat = cal.getTime(); //new Date();
        return new VBVariant(dateFormat.format(dat));
	}
	/*
	Returns a Date value.
	@author OlimilO
	@param  int year  year  of the date
	@param  int month month of the date
	@param  int day   day   of the date.
	*/	
	public static VBVariant DateSerial(int year, int month, int day){
	    DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		GregorianCalendar cal = new GregorianCalendar(year, month - 1, day);
        Date dat = cal.getTime(); 
        return new VBVariant(dateFormat.format(dat));
	}


	/*
	Gibt die Anzahl der Zeitintervalle zwischen zwei Datumsangaben zurück

	returns the amount of timespans between the two dates
	Interval: a string representing the kind of time
	Date1: one end of the timespan
	Date2: the other end of the timespan

	note: does not consider different first days of the week
	*/
	public static long DateDiff(String Interval, VBVariant Date1, VBVariant Date2)
	{
		return DateVal(Interval, Date2) - DateVal(Interval, Date1);
	}

	/* Helper function for DateDiff */
	private static long DateVal(String Interval, VBVariant d) {
		Date d1 = (Date)d.toObject();
		long sec = Conversion.CLng(VBVariant.valueOf(d1.getTime() * 0.001));
		int y = d1.getYear();
		long ret = 0;

		if (Interval.equals("s"))
			ret = sec;
		else if (Interval.equals("n"))
			ret = Conversion.CLng(VBVariant.valueOf(sec / 60));
		else if (Interval.equals("h"))
			ret = Conversion.CLng(VBVariant.valueOf(sec / 3600));
		else if (Interval.equals("d") || Interval.equals("y"))
			ret = Conversion.CLng(VBVariant.valueOf(sec / (24 * 3600)));
		else if (Interval.equals("w") || Interval.equals("ww"))
			ret = y * 52 + getWeekShort(d1);  // quick hack! not correct for multiple-year diffs
		else if (Interval.equals("m"))
			ret = 12 * y + d1.getMonth();
		else if (Interval.equals("q"))
			ret = 4 * y + d1.getMonth() / 4;
		else if (Interval.equals("yyyy"))
			ret = y;
		else {
			System.out.println("DateVal: unknown interval '" + Interval + "'");
			ret = 0;
		}

		return ret;
	}

	/* Helper function for DateDiff */
	private static int getWeekShort(Date d) {
		Calendar cal = Calendar.getInstance();

		cal.setTime(d);
		return cal.get(Calendar.WEEK_OF_YEAR);
        }

}
