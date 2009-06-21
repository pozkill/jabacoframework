package VBA;
import java.util.*;
import java.lang.*;

public class VBVariant implements java.lang.Comparable {
	Object Val = null;
	int iTyp = 0;
	
	public Class getValClass() {
		return Val.getClass();
	}	
	
	public Object getValObject() {
		return Val;
	}
	
	public VBVariant() {}
	public VBVariant(boolean i)			{ setVal(i ? 1 : 0); }
	public VBVariant(int i)    		 	{ setVal(i); }
	public VBVariant(double d)  		{ setVal(d); }
	public VBVariant(long l)    		{ setVal(l); }
	public VBVariant(Object o)  		{ setVal(o); }
	public VBVariant(java.util.Date d)  { setVal(d); }
	public VBVariant(String s)  		{ setVal(s); }
	public VBVariant(IVBArray a)  		{ setVal(a); }

	public boolean equals(VBVariant TestVal) {
		if (TestVal == null && Val == null) return true;
		if (Val == null) return false;
		if (TestVal == null) return false;
		try {
			if (Val == TestVal){
				return true;
			} else {
				if (TestVal.toString().equals(this.toString())) {
					return true;
				} else {
					return false;
				}
			}
		} catch (Exception e1) {
			return false;
		}
	}
	
	public int compareTo(VBVariant TestVal) throws ClassCastException {
		if ( TestVal == null ) return 1;
		if ( TestVal.toObject() == null && toObject() == null ) return 0;
		if ( TestVal.toObject() == null ) return 1;
		if ( toObject() == null ) return -1;
		if ( isNumeric() == false || TestVal.isNumeric() == false ) {
			if (isString() && TestVal.isString()) {
				return Strings.StrComp(toString(), TestVal.toString());
			} else if (toObject() instanceof java.lang.Comparable) {
				return ((java.lang.Comparable)toObject()).compareTo(TestVal.toObject()); 
			} else {
				return 1;
			} 
		}
		double d1 = doubleValue(); 
		double d2 = TestVal.doubleValue();
		if (d1 < d2) return -1;
		if (d1 > d2) return 1;
		return 0;
	}
	
	public int compareTo(Object TestVal) throws ClassCastException {
		if (TestVal instanceof VBVariant) {
			return compareTo((VBVariant)TestVal);
		} else if (toObject() instanceof java.lang.Comparable) {
			return ((java.lang.Comparable)toObject()).compareTo(TestVal);
		} else {
			throw new ClassCastException("Expected: VBVariant");
		}
	}
	
    public Iterator iterator() {
		if (toObject() instanceof VBArray) {
			return ((VBArray)toObject()).iterator();
		} else {
			ArrayList tmp = new ArrayList();
			tmp.add(toObject());
			return tmp.iterator();
		}
    }
	
	public static VBVariant valueOf(java.util.Date d) {
		return new VBVariant(d);
	}
	public static VBVariant valueOf(boolean b) {
		return new VBVariant(b?1:0);
	}
	public static VBVariant valueOf(float f) {
		return new VBVariant((double)f);
	}
	public static VBVariant valueOf(int i) {
		return new VBVariant(i);
	}
	public static VBVariant valueOf(double d) {
		return new VBVariant(d);
	}
	public static VBVariant valueOf(long l) {
		return new VBVariant(l);
	}
	public static VBVariant valueOf(Object o) {
		return new VBVariant(o);
	}
	public static VBVariant valueOf(String s) {
		return new VBVariant(s);
	}
	public static VBVariant valueOf(IVBArray v) {
		return new VBVariant(v);
	}
	
	private String withoutFloating(String val) {
		Double tmpDbl = new Double(val);
		Long tmpLng = new Long(tmpDbl.longValue());
		return tmpLng.toString();	
	}
	
	public boolean isString() {
		try {
			if (toObject() instanceof java.lang.String) return true;	
		} catch(Exception e) {}
		return false;
	}
	
	public boolean isDate() {
		try {
			if (toObject() instanceof java.util.Date) return true;	
		} catch(Exception e) {}
		return false;
	}
	
	public java.util.Date toDate() {
		return ((java.util.Date)toObject());
	}
	
	
	private Object getVal(int iReturnTyp) {
		switch (iTyp) {
			case 1: // Double
				switch (iReturnTyp) {
					case 1: // Double
						return Val;
					case 2: // Integer
						return (new Integer((int)((Double)Val).doubleValue()));
					case 3: // Long
						return (new Long((long)((Double)Val).doubleValue()));
					case 4: // String
						double tmpValDbl = ((Double)Val).doubleValue();
						long tmpValLng = ((Double)Val).longValue();
						if ((tmpValDbl % tmpValLng) == 0) {
							return String.valueOf(tmpValLng);
						} else {
							return String.valueOf(tmpValDbl);
						}
				}
			break;
			case 2: // Integer
				switch (iReturnTyp) {
					case 1: // Double
						return (new Double((double)((Integer)Val).intValue()));
					case 2: // Integer
						return Val;
					case 3: // Long
						return (new Long((long)((Integer)Val).intValue()));
					case 4: // String
						return ((Integer)Val).toString();
				}
			break;
			case 3: // Long
				switch (iReturnTyp) {
					case 1: // Double
						return (new Double((double)((Long)Val).longValue()));
					case 2: // Integer
						return (new Integer((int)((Long)Val).longValue()));
					case 3: // Long
						return Val;
					case 4: // String
						return ((Long)Val).toString();
				}
			break;
			case 4: // String
				switch (iReturnTyp) {
					case 1: // Double
						return new Double((String)Val);
					case 2: // Integer
						return new Integer(withoutFloating((String)Val));
					case 3: // Long
						return new Long(withoutFloating((String)Val));
					case 4: // String
						return Val;
				}
			break;
			case 5: // object
				if (iReturnTyp > 0 && iReturnTyp < 5) {
					switch (iReturnTyp) {
						case 1: // Double
							return (Double.valueOf(Val.toString()));
						case 2: // Integer
							return (Integer.valueOf(Val.toString()));
						case 3: // Long
							return (Long.valueOf(Val.toString()));
						case 4: // String
							return Val.toString();
					}
				}
			break;
		}

		return Val;
	}
	
	
	public double doubleValue() {
		try {
			return ((Double)getVal(1)).doubleValue();
		} catch (Exception e) { return 0.0; }
	}
	
	public int intValue() {
		try {
			return ((Integer)getVal(2)).intValue();
		} catch (Exception e) { return 0; }
	}
	
	public long longValue() {
		try {
			return ((Long)getVal(3)).longValue();
		} catch (Exception e) { return 0; }
	}
	
	public boolean isNumeric() {
		boolean ret = (this.doubleValue() != 0);
		if (ret == false) {
			switch (iTyp) {
				case 1: // Double
				case 2: // Integer
				case 3: // Long
					ret = true;
			}
		}
		return ret;
	}
	
	public Object toObject() {
		try {
			return ((Object)getVal(5));
		} catch (Exception e) { return null; }
	}
	
	public String toString() {
		if (isDate()) return (Conversion.CStr(toDate()));
		try {
			if (getVal(4) == null)
				if (Val != null) {
					return (Val.toString());
				} else {
					return ("");
				}
			else
				return (String)getVal(4);
		} catch (Exception e) { return (""); }
	}
	
	private void setVal(double d) {
		Val = new Double(d); iTyp = 1;
	}
	private void setVal(int i) {
		Val = new Integer(i); iTyp = 2;
	}
	private void setVal(long l) {
		Val = new Long(l); iTyp = 3;
	}
	private void setVal(String s) {
		Val = (String)s; iTyp = 4;
	}
	private void setVal(Object o) {
		if (o instanceof VBEnumClass) {
			setVal(((VBEnumClass)o).longValue());
		} else if (o instanceof java.lang.Double) {
			setVal(((java.lang.Double)o).doubleValue());
		} else if (o instanceof java.lang.Integer) {
			setVal(((java.lang.Integer)o).intValue());
		} else if (o instanceof java.lang.Long) {
			setVal(((java.lang.Long)o).longValue());
		} else if (o instanceof java.lang.String) {
			setVal((java.lang.String)o);
		} else if (o instanceof java.lang.Boolean) {
			setVal(((java.lang.Boolean)o).booleanValue() ? 1 : 0);
		} else if (o instanceof java.lang.Byte) {
			setVal(((java.lang.Byte)o).intValue());
		} else {
			Val = o; iTyp = 5;
		}
	}
}
