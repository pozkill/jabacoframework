package VBA;
import java.util.*;
import java.lang.*;

public class VBVariant implements java.lang.Comparable {
	Object Val = null;
/*
   vbEmpty           = 0
   vbNull            = 1
   vbInteger         = 2
   vbLong            = 3
   vbSingle          = 4
   vbDouble          = 5
   vbCurrency        = 6
   vbDate            = 7
   vbString          = 8
   vbObject          = 9
   vbError           = 10
   vbBoolean         = 11
   vbVariant         = 12
   vbDataObject      = 13
   vbDecimal         = 14
   vbByte            = 17
   vbUserDefinedType = 36
   vbArray           = 8192
*/
	int iTyp = 0;
	
	public Class getValClass() {
		return Val.getClass();
	}	
	
	public Object getValObject() {
		return Val;
	}
	
	public VBVariant() 					{ iTyp =  0; } //iTyp =  0
	public VBVariant(int i)    		 	{ setVal(i); } //iTyp =  2
	public VBVariant(long l)    		{ setVal(l); } //iTyp =  3
	public VBVariant(float f)    		{ setVal(f); } //iTyp =  4
	public VBVariant(double d)  		{ setVal(d); } //iTyp =  5
	//public VBVariant(currency c)  	{ setVal(c); } //iTyp =  6
	public VBVariant(java.util.Date d)  { setVal(d); } //iTyp =  7
	public VBVariant(String s)  		{ setVal(s); } //iTyp =  8
	public VBVariant(Object o)  		{ setVal(o); } //iTyp =  9
	public VBVariant(Throwable e)		{ setVal(e); } //iTyp = 10
	public VBVariant(boolean b)			{ setVal(b); } //iTyp = 11
	//public VBVariant(VBVariant v)		{ setVal(v); } //iTyp = 12
	//public VBVariant(DataObject data)	{ setVal(data); } //iTyp = 13
	//public VBVariant(Decimal decim)	{ setVal(decim); } //iTyp = 14
	public VBVariant(byte b)    		{ setVal(b); } //iTyp = 17
	public VBVariant(VBTypeClass udt)	{ setVal(udt); } //iTyp = 36
	public VBVariant(IVBArray a)  		{ setVal(a); }   //iTyp = 8192 + VarType(ArrayElem)
	
	public int VType() {
	   return iTyp;
	}
	public int innerVarType() throws ClassCastException  {
		if (iTyp == 12) {
			VBVariant var = (VBVariant)Val;
			return var.innerVarType();
		}
		else {
			return iTyp;
		}
	}
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
	
	public static VBVariant valueOf(int i) {
		return new VBVariant(i);
	}
	public static VBVariant valueOf(long l) {
		return new VBVariant(l);
	}
	public static VBVariant valueOf(float f) {
		return new VBVariant(f);
	}
	public static VBVariant valueOf(double d) {
		return new VBVariant(d);
	}
	public static VBVariant valueOf(java.util.Date d) {
		return new VBVariant(d);
	}
	public static VBVariant valueOf(boolean b) {
		return new VBVariant(b); //?1:0);
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
	
	/*
	private String getHexString() {
		String h = (String)Val;
		if ((h.startswith("&H") | h.startswith("0x")) & (h.length <= 18)) {
			h = h.subString(2, 16);
		}
		else {h = null;}
		return h;
	}
	private Long parseHex() {
		long ret = 0;
		if isHexString() {
			ret = java.lang.Long.valueOf(
		}
		return ret;
	}
	*/
	/*
	   vbEmpty           = 0
   vbNull            = 1
   vbInteger         = 2
   vbLong            = 3
   vbSingle          = 4
   vbDouble          = 5
   vbCurrency        = 6
   vbDate            = 7
   vbString          = 8
   vbObject          = 9
   vbError           = 10
   vbBoolean         = 11
   vbVariant         = 12
   vbDataObject      = 13
   vbDecimal         = 14
   vbByte            = 17
   vbUserDefinedType = 36
   vbArray           = 8192
*/
	private Object getVal(int iReturnTyp) {
		switch (iTyp) {
		//ummm no, why?
		//Double must be 5, string must be 8
			case 2: // Integer
				switch (iReturnTyp) {
					case 2: // Integer
						return Val;
					case 3: // Long
						return (new Long((long)((Integer)Val).intValue()));
					case 5: // Double
						return (new Double((double)((Integer)Val).intValue()));
					case 8: // String
						return ((Integer)Val).toString();
				}
			break;
			case 3: // Long
				switch (iReturnTyp) {
					case 2: // Integer
						return (new Integer((int)((Long)Val).longValue()));
					case 3: // Long
						return Val;
					case 5: // Double
						return (new Double((double)((Long)Val).longValue()));
					case 8: // String
						return ((Long)Val).toString();
				}
			break;
/*			case 4: // Single
				switch (iReturnTyp) {
					case 2: // Integer
						return (new Integer((int)((Double)Val).doubleValue()));
					case 3: // Long
						return (new Long((long)((Double)Val).doubleValue()));
					case 4: // Single
						return (float)Val;
					case 5: // Double
						return Val;
					case 8: // String
						double tmpValDbl = ((Double)Val).doubleValue();
						long tmpValLng = ((Double)Val).longValue();
						if ((tmpValDbl % tmpValLng) == 0) {
							return String.valueOf(tmpValLng);
						} else {
							return String.valueOf(tmpValDbl);
						}
				}
			break;*/
			case 5: // Double
				switch (iReturnTyp) {
					case 2: // Integer
						return (new Integer((int)((Double)Val).doubleValue()));
					case 3: // Long
						return (new Long((long)((Double)Val).doubleValue()));
					case 4: // Single
						return (Float)Val;
					case 5: // Double
						return Val;
					case 8: // String
						double tmpValDbl = ((Double)Val).doubleValue();
						long tmpValLng = ((Double)Val).longValue();
						if ((tmpValDbl % tmpValLng) == 0) {
							return String.valueOf(tmpValLng);
						} else {
							return String.valueOf(tmpValDbl);
						}
				}
			break;
			//case 6: //Currency
			//case 7: //Date
			case 8: // String
				switch (iReturnTyp) {
					case 2: // Integer
						if ( (((String)Val).trim().length()>2) &&  (
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&H")  ||
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&O")  ||
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&B") ) ) {
							return new Integer( (int) new VBVariant(Val).longValue() );
						} else {
							return new Integer(withoutFloating((String)Val));
						}
					case 3: // Long
						if ((((String)Val).trim().length()>2) && (((String)Val).trim().substring(0,2).equalsIgnoreCase("&H"))) {
							String str = ((String)Val).trim().substring(2);
							int pos=0;
							long p=0, v=0;
							for (pos = 0; pos<=str.length()-1; pos++) {
								v = "0123456789abcdef".indexOf(str.substring(pos,pos+1));
								if (v<0) return 0;
								p = (16 * p) + v;
							}
							return new Long(p);
						} else if ((((String)Val).trim().length()>2) && (((String)Val).trim().substring(0,2).equalsIgnoreCase("&O"))) {
							String str = ((String)Val).trim().substring(2);
							int pos=0;
							long p=0, v=0;
							for (pos = 0; pos<=str.length()-1; pos++) {
								v = "01234567".indexOf(str.substring(pos,pos+1));
								if (v<0) return 0;
								p = (8 * p) + v;
							}
							return new Long(p);
						} else if ((((String)Val).trim().length()>2) && (((String)Val).trim().substring(0,2).equalsIgnoreCase("&B"))) {
							String str = ((String)Val).trim().substring(2);
							int pos=0;
							long p=0, v=0;
							for (pos = 0; pos<=str.length()-1; pos++) {
								v = "01".indexOf(str.substring(pos,pos+1));
								if (v<0) return 0;
								p = (2 * p) + v;
							}
							return new Long(p);
						} else {
							return new Long(withoutFloating((String)Val));
						}
					case 5: // Double
						if ( (((String)Val).trim().length()>2) &&  (
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&H")  ||
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&O")  ||
						    ((String)Val).trim().substring(0,2).equalsIgnoreCase("&B") ) ) {
							return new Double( new VBVariant(Val).longValue() );
						} else {
							return new Double((String)Val);
						}
					case 8: // String
						return Val;
				}
			break;
			case 9: // object
				if (iReturnTyp > 0 && iReturnTyp < 5) {
					switch (iReturnTyp) {
						case 2: // Integer
							return (Integer.valueOf(Val.toString()));
						case 3: // Long
							return (Long.valueOf(Val.toString()));
						case 5: // Double
							return (Double.valueOf(Val.toString()));
						case 8: // String
							return Val.toString();
					}
				}
			break;
			case 11: //Boolean
				switch (iReturnTyp) {
					case 2: // Integer
						return (Integer.valueOf(Val.toString()));
					case 3: // Long
						return (Long.valueOf(Val.toString()));
					case 5: // Double
						return (Double.valueOf(Val.toString()));
					case 8: // String
						return Val.toString();
					//case 9: //Object
					//	return Val;
					case 11: // Boolean
						return Val;
				}
			break;
				
		}
		return Val;
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
	/*public long singleValue() {
		try {
			return ((Long)getVal(4)).longValue();
		} catch (Exception e) { return 0; }
	}*/
	
	public double doubleValue() {
		try {
			return ((Double)getVal(5)).doubleValue();
		} catch (Exception e) { return 0.0; }
	}
	
	public boolean isNumeric() {
		boolean ret = (this.doubleValue() != 0);
		if (ret == false) {
			switch (iTyp) {
				case 2: // Integer
				case 3: // Long
				case 4: // Single (float)
				case 5: // Double
				case 6: // Currency
				        // Decimal
					ret = true;
			}
		}
		return ret;
	}
	
	public Object toObject() {
		try {
			return ((Object)getVal(9));
		} catch (Exception e) { return null; }
	}
	
	public String toString() {
		if (isDate()) return (Conversion.CStr(toDate()));
		try {
			if (getVal(8) == null)
				if (Val != null) {
					return (Val.toString());
				} else {
					return ("");
				}
			else
				return (String)getVal(8);
		} catch (Exception e) { return (""); }
	}
	//****************\\  setVal  //****************\\
	public void setVal(int i) {
		Val = new Integer(i); iTyp = 2;
	}
	public void setVal(long l) {
		Val = new Long(l);    iTyp = 3;
	}
	public void setVal(float f) {
		Val = new Float(f);    iTyp = 4;
	}
	public void setVal(double d) {
		Val = new Double(d);   iTyp = 5;
	}
/*	public void setVal(VBCurrency c) {
		Val = new VBCurrency(c);  iTyp = 6;
	}*/
	public void setVal(java.util.Date d) {
		Val = d;  iTyp = 7;
	}
	public void setVal(String s) {
		Val = (String)s;      iTyp = 8;
	}
	public void setVal(Object o) {
		       if (o instanceof VBEnumClass) {
			setVal(((VBEnumClass)o).longValue());
		} else if (o instanceof VBTypeClass) {
			setVal(((VBTypeClass)o));//.intValue());
		} else if (o instanceof java.lang.Double) {
			setVal(((java.lang.Double)o).doubleValue());
		} else if (o instanceof java.lang.Integer) {
			setVal(((java.lang.Integer)o).intValue());
		} else if (o instanceof java.lang.Long) {
			setVal(((java.lang.Long)o).longValue());
		} else if (o instanceof java.lang.String) {
			setVal((java.lang.String)o);
		} else if (o instanceof java.lang.Boolean) {
			setVal(((java.lang.Boolean)o).booleanValue()); // ? 1 : 0);
		} else if (o instanceof java.lang.Byte) {
			setVal(((java.lang.Byte)o));//.intValue());
		} else if (o instanceof java.lang.Throwable) {
			setVal(((java.lang.Throwable)o));//.intValue());
		} else {
			Val = o; iTyp = 9; //vbObject
		}
	}
	public void setVal(Throwable e) {
		Val = e; iTyp = 10;
	}
	
	public void setVal(boolean b) {
		Val = new Boolean(b); iTyp = 11;
	}
	public void setVal(VBVariant var) {
		Val = new VBVariant(var); iTyp = 12;
	}
    //vbDataObject      = 13
    //vbDecimal         = 14
	public void setVal(byte b) {
		Val = new Byte(b); iTyp = 17;
	}
    //vbUserDefinedType = 36
	public void setVal(VBTypeClass udt) {
		Val = udt; iTyp = 36;
	}
    //vbArray           = 8192	
	public void setVal(IVBArray a) {
		Val = a; iTyp = 8192;
		       if (a instanceof VBArrayInteger) {
			iTyp += 2;
		} else if (a instanceof VBArrayLong) {
			iTyp += 3;
		} else if (a instanceof VBArraySingle) {
			iTyp += 4;
		} else if (a instanceof VBArrayDouble) {
			iTyp += 5;
		} else if (a instanceof VBArrayDate) {
			iTyp += 7;
		} else if (a instanceof VBArrayString) {
			iTyp += 8;
		} else if (a instanceof VBArrayObject) {
			iTyp += 9;
		} else if (a instanceof VBArrayBoolean) {
			iTyp += 11;
		} else if (a instanceof VBArrayVariant) {
			iTyp += 12;
		} else if (a instanceof VBArrayByte) {
			iTyp += 17;
		}
	}
}
