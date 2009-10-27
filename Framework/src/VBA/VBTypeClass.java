package VBA;
import java.io.*;

public class VBTypeClass implements Cloneable, java.io.Serializable {
	
	// just flagging
	public Object clone() {
		try {
			return super.clone();
		} catch (Exception e) {
			System.out.println (e.toString());
			//throw new InternalError(e.toString());
			return null;
		}
	}
	
	private void writeObject(ObjectOutputStream out) throws IOException {
		out.defaultWriteObject();
		//2009-10-21 OlimilO todo: implement VBPut
	}
	private void readObject(ObjectInputStream in) throws IOException, ClassNotFoundException {
		in.defaultReadObject();
		//2009-10-21 OlimilO todo: implement VBGet
	}
	/* better move VBGet, VBPut to FileSystem
	 * so it's better possible in Jabaco-code */
	private void VBGet(DataInputStream Instream) throws Exception {
		java.lang.reflect.Field f;
		java.lang.reflect.Field[] fs = this.getClass().getFields();
		for (int i = 0; i < fs.length; ++i) {
			f = fs[i];			
			if (f.getType().isPrimitive() == true) {
				String n = f.getType().getName().toLowerCase();
				       if (n == "byte")   { f.setByte   (this, Instream.readByte());
				} else if (n == "char")   { f.setByte   (this, Instream.readByte());
				} else if (n == "short")  { f.setShort  (this, Instream.readShort());
				} else if (n == "int")    { f.setInt    (this, Instream.readShort());
				} else if (n == "long")   { f.setLong   (this, Instream.readInt());
				} else if (n == "float")  { f.setFloat  (this, Instream.readFloat());
				} else if (n == "double") { f.setDouble (this, Instream.readDouble());
				}
			}
			else{
				//to do: implement Get for Arrays, objects   
			}
		}
	}//*/
	/* better move VBGet, VBPut to FileSystem
	 * so it's better possible in Jabaco-code */
	private void VBPut(DataOutputStream Outstream) throws Exception {
		java.lang.reflect.Field f;
		java.lang.reflect.Field[] fs = this.getClass().getFields();
		for (int i = 0; i < fs.length; ++i) {
			f = fs[i];			
			if (f.getType().isPrimitive() == true) {
				String n = f.getType().getName().toLowerCase();
				       if (n == "byte")   { Outstream.writeByte   (f.getByte(this));
				} else if (n == "char")   { Outstream.writeChar   (f.getByte(this));   
				} else if (n == "short")  { Outstream.writeShort  (f.getShort(this));  
				} else if (n == "int")    { Outstream.writeShort  (f.getInt(this));    
				} else if (n == "long")   { Outstream.writeInt    ((int)f.getLong(this));   
				} else if (n == "float")  { Outstream.writeFloat  (f.getFloat(this));  
				} else if (n == "double") { Outstream.writeDouble (f.getDouble(this)); 
				}
			}
			else{
				//to do: implement Get for Arrays, objects   
			}
		}
	}//*/
}
