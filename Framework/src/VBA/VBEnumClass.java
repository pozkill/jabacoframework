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
	
}
