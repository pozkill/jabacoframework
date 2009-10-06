package VBA;
import java.lang.*;
import java.lang.Math;

public class Financial {

/*
Function DDB(Cost As Double, Salvage As Double, Life As Double, Period As Double, [Factor]) As Double
Function FV(Rate As Double, NPer As Double, Pmt As Double, [PV], [Due]) As Double
Function IPmt(Rate As Double, Per As Double, NPer As Double, PV As Double, [FV], [Due]) As Double
Function IRR(ValueArray() As Double, [Guess]) As Double
Function MIRR(ValueArray() As Double, FinanceRate As Double, ReinvestRate As Double) As Double
Function NPer(Rate As Double, Pmt As Double, PV As Double, [FV], [Due]) As Double
Function NPV(Rate As Double, ValueArray() As Double) As Double
Function Pmt(Rate As Double, NPer As Double, PV As Double, [FV], [Due]) As Double
Function PPmt(Rate As Double, Per As Double, NPer As Double, PV As Double, [FV], [Due]) As Double
Function PV(Rate As Double, NPer As Double, Pmt As Double, [FV], [Due]) As Double
Function Rate(NPer As Double, Pmt As Double, PV As Double, [FV], [Due], [Guess]) As Double
Function SLN(Cost As Double, Salvage As Double, Life As Double) As Double
Function SYD(Cost As Double, Salvage As Double, Life As Double, Period As Double) As Double
*/
	public static double DDB(double Cost, double Salvage, double Life, double Period) {
		return DDB(Cost, Salvage, Life, Period, 2.0);
	}

	public static double DDB(double Cost, double Salvage, double Life, double Period, double Factor) {
		if ((((Factor <= 0.0) || (Salvage < 0.0)) || (Period <= 0.0)) || (Period > Life))
		{
			//throw new ArgumentException(Utils.GetResourceString("Argument_InvalidValue1", new string[] { "Factor" }));
		}
		if (Cost > 0.0)
		{
			double num4;
			double num5;
			if (Life < 2.0)
			{
				return (Cost - Salvage);
			}
			if ((Life == 2.0) && (Period > 1.0))
			{
				return 0.0;
			}
			if ((Life == 2.0) && (Period <= 1.0))
			{
				return (Cost - Salvage);
			}
			if (Period <= 1.0)
			{
				num4 = (Cost * Factor) / Life;
				num5 = Cost - Salvage;
				if (num4 > num5)
				{
					return num5;
				}
				return num4;
			}
			num5 = (Life - Factor) / Life;
			double y = Period - 1.0;
			num4 = ((Factor * Cost) / Life) * Math.pow(num5, y);
			double num6 = Cost * (1.0 - Math.pow(num5, Period));
			double num2 = (num6 - Cost) + Salvage;
			if (num2 > 0.0)
			{
				num4 -= num2;
			}
			if (num4 >= 0.0)
			{
				return num4;
			}
		}
		return 0.0;
	}


}
