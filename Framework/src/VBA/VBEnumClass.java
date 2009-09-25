package VBA;

public class VBEnumClass extends VBVariant implements IResource {
	

	// just flagging
	public VBEnumClass() 		  { super(0); }
	public VBEnumClass(int i)     { super(i); }
	public VBEnumClass(long l)    { super(l); }
	public VBEnumClass(String s)  { super(s); }
	/*
	public boolean equals(VBVariant o) {
		if (o == null) return false;
		try {
			if (this == o){
				return true;
			} else {
				if (o.toString().trim().toLowerCase().equals(this.toString().trim().toLowerCase())) {
					return true;
				} else {
					return false;
				}
			}
		} catch (Exception e1) {
			return false;
		}
	}*/
	/**
	 * returns array of names of the constant static fields of the enum
	 */
	public String[] getEnumNames() {
	       java.lang.reflect.Field[] fs = this.getClass().getFields();
	       String[] s = new String[fs.length];
	       for (int i = 0; i < fs.length; ++i) {
               s[i] = fs[i].getName();
           }
	       return s;
    }
	/**
	 * returns array of names of the constant static fields of the enum
	 */
	public VBVariant[] getEnumValues() throws Exception {
	       java.lang.reflect.Field[] fs = this.getClass().getFields();
	       VBVariant[] v = new VBVariant[fs.length];
	       for (int i = 0; i < fs.length; ++i) {
               v[i] = (VBVariant)fs[i].get(this);
           }
	       return v;
    }	
}
