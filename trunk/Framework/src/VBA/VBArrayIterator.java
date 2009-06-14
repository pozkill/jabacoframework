package VBA;
import java.util.*;

public class VBArrayIterator implements Iterator {
	int cursor = 0;
	Object[] colObject;
	
	public VBArrayIterator(Object[] refObject) {
		colObject = refObject;
	}
	
	public boolean hasNext() {
		return cursor < colObject.length;
	}

	public Object next() {
		Object next = colObject[cursor];
		cursor++;
		return next;
	}

	public void remove() {
		cursor++;
	}
	
}

