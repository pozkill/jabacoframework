package VBA;

import java.lang.*;
	/**
	 * @author Manuel Siekmann
	 * @version 1.0 @since 2009-09-05
	 */
public class Math {
	/**
	 * Returns the sine of an angle. 
	 * @param  Number can be any valid numeric expression that expresses an angle in radians. 
	 */
	public static double Sin(double Number) {
		return java.lang.Math.sin(Number);
	}	
	/**
	 * Returns the cosine of an angle.
	 * @author Manuel Siekmann
	 * @param  Number can be any valid numeric expression that expresses an angle in radians. 
	 */
	public static double Cos(double Number) {
		return java.lang.Math.cos(Number);
	}	
	/**
	 * Returns the tangent of an angle. 
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression that expresses an angle in radians. 
	 */
	public static double Tan(double Number) {
		return java.lang.Math.tan(Number);
	}
	/**
	 * Returns the cosecans of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Csc(double Number) {
		return 1 / java.lang.Math.sin(Number);
	}
	/**
	 * Returns the secant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Sec(double Number) {
		return 1 / java.lang.Math.cos(Number);
	}
	/**
	 * Returns the cotangent of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Cot(double Number) {
		return 1 / java.lang.Math.tan(Number);
	}

	/**
	 * Returns the arc sine of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Asin(double Number) {
		return java.lang.Math.asin(Number);
	}
	/**
	 * Returns the arc cosine of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Acos(double Number) {
		return java.lang.Math.acos(Number);
	}
	/**
	 * Returns the arc tangent of a number.
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression.
	 */
	public static double Atn(double Number) {
		return java.lang.Math.atan(Number);
	}
	/**
	 * Returns the arc tangent of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Atan2(double y, double x) {
		return java.lang.Math.atan2(y, x);
	}

	/**
	 * Returns the arc cosecant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Acsc(double y) {
		return java.lang.Math.asin(1 / y);
	}

	/**
	 * Returns the arc secant of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Asec(double x) {
		return java.lang.Math.acos(1 / x);
	}

	/**
	 * Returns the arc cotangent of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Acot(double t) {
		return java.lang.Math.PI * 0.5 - java.lang.Math.atn(t);
	}

	/**
	 * Returns the hyperbolic sine of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Sinh(double Number) {
		return java.lang.Math.sinh(Number);
	}

	/**
	 * Returns the hyperbolic cosine of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Cosh(double Number) {
		return java.lang.Math.cosh(Number);
	}

	/**
	 * Returns the hyperbolic tangent of a number.
	 * @author OlimilO
	 * @param  The number arguments can be any valid numeric expressions.
	 * @since  2009-09-05
	 */
	public static double Tanh(double Number) {
		return java.lang.Math.tanh(Number);
	}

	/**
	 * Returns the hyperbolic cosecant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Csch(double Number) {
        return 2 / (java.lang.Math.exp(Number) - java.lang.Math.exp(-Number));
	}

	/**
	 * Returns the hyperbolic secant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Sech(double Number) {
        return 2 / (java.lang.Math.exp(Number) + java.lang.Math.exp(-Number));
	}

	/**
	 * Returns the hyperbolic cotangent of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Coth(double Number) {
	    //(Exp(t) + Exp(-t)) / (Exp(t) - Exp(-t))
        return ((java.lang.Math.exp(Number) + java.lang.Math.exp(-Number)) / ((java.lang.Math.exp(Number) - java.lang.Math.exp(-Number));
	}

	/**
	 * Returns the area hyperbolic sine of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Asinh(double Number) {
	    //Log(y + Sqr(y * y + 1))
		return java.lang.Math.log(Number + java.lang.Math.sqrt(Number * Number + 1));
	}

	/**
	 * Returns the area hyperbolic cosine of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Acosh(double Number) {
	    //Log(x + Sqr(x * x - 1))
		return java.lang.Math.log(Number + java.lang.Math.sqrt(Number * Number - 1));
	}

	/**
	 * Returns the area hyperbolic tangent of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Atanh(double Number) {
	    //Log((1 + t) / (1 - t)) / 2
		return java.lang.Math.log((1 + Number) / (1 - Number)) / 2;
	}

	/**
	 * Returns the area hyperbolic cosecant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Acsch(double Number) {
	    //Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
		return java.lang.Math.log((Sgn(Number) * java.lang.Math.sqrt(Number * Number + 1) + 1) / Number);
	}

	/**
	 * Returns the area hyperbolic secant of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Asech(double Number) {
	    //Log((Sqr(-x * x + 1) + 1) / x)
		return java.lang.Math.log((java.lang.Math.sqrt(-Number * Number + 1) + 1) / Number);
	}

	/**
	 * Returns the area hyperbolic cotangent of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double Acoth(double Number) {
	    //Log((x + 1) / (x - 1)) / 2
		return java.lang.Math.log((Number + 1) / (Number - 1)) / 2;
	}

	/**
	 * Returns the sine cardinale of a number.
	 * @author OlimilO
	 * @param  The number argument can be any valid numeric expression.
	 * @since  2009-09-05
	 */
	public static double SinC(double Number) {
		return (Number == 0.0) ? 1 : (java.lang.Math.sin(Number) / Number);
	}

	/**
	 * Returns the absolute value of a number.
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression.
	 */
	public static double Abs(double Number) {
		return java.lang.Math.abs(Number);
	}
	/**
	 * Returns e (the base of natural logarithms) raised to a power.
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression. 
	 */
	public static double Exp(double Number) {
		return java.lang.Math.exp(Number);
	}
	/**
	 * Returns the natural logarithm of a number.
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression. 
	 */
	public static double Log(double Number) {
		return java.lang.Math.log(Number);
	}
	/* Unneeded */
	public static void Randomize() {/**/}
	/* Unneeded */
	public static void Randomize(double Number) {/**/}
	/**
	 * Returns a random double value with a positive sign, greater than or equal to 0.0 and less than 1.0.
	 * @author Manuel Siekmann
	 */
	public static double Rnd() {
		return Rnd(0);
	}
	/**
	 * Returns a random double value with a positive sign, greater than or equal to 0.0 and less than 1.0.
	 * @author Manuel Siekmann
	 * @param  unused
	 */
	public static double Rnd(double Number) {
		return java.lang.Math.random();
	}
	/**
	 * Returns a number rounded to a specified number of decimal places. 
	 * @author Manuel Siekmann
	 * @param  Numeric expression being rounded. 
	 */
	public static double Round(double Number) {
		return Round(Number, 0);
	}
	/**
	 * Returns a number rounded to a specified number of decimal places. 
	 * @author Manuel Siekmann
	 * @param  Numeric expression being rounded. 
	 * @param  Number indicating how many places to the right of the decimal are included in the rounding. If omitted, integers are returned by the Round function. 
	 */
	public static double Round(double Number, int NumDigitsAfterDecimal) {
		double faktor = java.lang.Math.pow(10, NumDigitsAfterDecimal);
		return (java.lang.Math.round(Number * faktor) / faktor);
	}
	/**
	 * Returns an integer indicating the sign of a number. 
	 * @return  The Sgn function has the following return values:
	<div class="tableSection">
	<tr><th>If number is</th><th>Sgn returns</th></tr>	
	<tr><td>Greater than zero</td><td>1</td></tr>
	<tr><td>Equal to zero</td><td>0</td></tr>
	<tr><td>Less than zero</td><td>-1</td></tr>
	</table>
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression.
	 */
	public static double Sgn(double Number) {
		if (Number > 0) return 1;
		if (Number < 0) return -1;
		return 0;
		//return java.lang.Math.signum(Number); not in Java 1.4.2
	}
	/**
	 * Returns the square root of a number.
	 * @author Manuel Siekmann
	 * @param  The number argument can be any valid numeric expression greater than or equal to 0. 
	 */
	public static double Sqr(double Number) {
		return java.lang.Math.sqrt(Number);
	}
}
