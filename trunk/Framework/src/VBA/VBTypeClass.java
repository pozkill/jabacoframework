package VBA;

public class VBTypeClass implements Cloneable {
	
	public Object clone() {
		try {
			return super.clone();
		} catch (Exception e) {
			System.out.println (e.toString());
			//throw new InternalError(e.toString());
			return null;
		}
	}
	
}
