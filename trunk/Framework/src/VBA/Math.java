package VBA;

import java.lang.*;

public class Math {
	/*
	Returns the absolute value of a number.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression.
	*/
	public static double Abs(double Number) {
		return java.lang.Math.abs(Number);
	}
	/*
	Returns the arctangent of a number.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression.
	*/
	public static double Atn(double Number) {
		return java.lang.Math.atan(Number);
	}
	/*
	Returns the cosine of an angle.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression that expresses an angle in radians. 
	*/
	public static double Cos(double Number) {
		return java.lang.Math.cos(Number);
	}
	/*
	Returns e (the base of natural logarithms) raised to a power.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression. 
	*/
	public static double Exp(double Number) {
		return java.lang.Math.exp(Number);
	}
	/*
	Returns the natural logarithm of a number.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression. 
	*/
	public static double Log(double Number) {
		return java.lang.Math.log(Number);
	}
	/* Unneeded */
	public static void Randomize() {/**/}
	/* Unneeded */
	public static void Randomize(double Number) {/**/}
	/*
	Returns a random double value with a positive sign, greater than or equal to 0.0 and less than 1.0.
	@author Manuel Siekmann
	*/
	public static double Rnd() {
		return Rnd(0);
	}
	/*
	Returns a random double value with a positive sign, greater than or equal to 0.0 and less than 1.0.
	@author Manuel Siekmann
	@param  unused
	*/
	public static double Rnd(double Number) {
		return java.lang.Math.random();
	}
	/*
	Returns a number rounded to a specified number of decimal places. 
	@author Manuel Siekmann
	@param  Numeric expression being rounded. 
	*/
	public static double Round(double Number) {
		return Round(Number, 0);
	}
	/*
	Returns a number rounded to a specified number of decimal places. 
	@author Manuel Siekmann
	@param  Numeric expression being rounded. 
	@param  Number indicating how many places to the right of the decimal are included in the rounding. If omitted, integers are returned by the Round function. 
	*/
	public static double Round(double Number, int NumDigitsAfterDecimal) {
		double faktor = java.lang.Math.pow(10, NumDigitsAfterDecimal);
		return (java.lang.Math.round(Number * faktor) / faktor);
	}
	/*
	Returns an integer indicating the sign of a number. 
	@return  The Sgn function has the following return values:
	<div class="tableSection">
	<tr><th>If number is</th><th>Sgn returns</th></tr>	
	<tr><td>Greater than zero</td><td>1</td></tr>
	<tr><td>Equal to zero</td><td>0</td></tr>
	<tr><td>Less than zero</td><td>-1</td></tr>
	</table>
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression.
	*/
	public static double Sgn(double Number) {
		if (Number > 0) return 1;
		if (Number < 0) return -1;
		return 0;
	}
	/*
	Returns the sine of an angle. 
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression that expresses an angle in radians. 
	*/
	public static double Sin(double Number) {
		return java.lang.Math.sin(Number);
	}
	/*
	Returns the square root of a number.
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression greater than or equal to 0. 
	*/
	public static double Sqr(double Number) {
		return java.lang.Math.sqrt(Number);
	}
	/*
	Returns the tangent of an angle. 
	@author Manuel Siekmann
	@param  The number argument can be any valid numeric expression that expresses an angle in radians. 
	*/
	public static double Tan(double Number) {
		return java.lang.Math.tan(Number);
	}
}
