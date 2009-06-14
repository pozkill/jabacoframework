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
		return new VBVariant(cal.get(Calendar.MONTH));
	}
	/*
	Returns an Integer value from 1 through 12 representing the month of the year.
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
        Date date = new Date();
        return new VBVariant(dateFormat.format(date));
	}
	
}
