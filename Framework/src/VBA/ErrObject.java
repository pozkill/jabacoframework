package VBA;
import java.lang.*;
import java.io.PrintStream;

/**
 * @author OlimilO
 * @since 2009_10_06
 */
public class ErrObject {
	//extends Exception???
	//extends PrintStream???
	private PrintStream myErr = null;
	private int myNumber = 0;
	private String mySource = null;
	private String myDescription = null;
	private String myHelpFile = null;
	private String myHelpContext = null;
	
	public ErrObject() {
		//myErr = java.lang.System.err();
	}
	public void Clear() {
		myNumber = 0;
		mySource = null;
		myDescription = null;
		myHelpFile = null;
		myHelpContext = null;
	}
	public void $Description(String newDescription) {
		myDescription = newDescription;
	}
	public String $Description() {
		return myDescription;
	}
	public void $HelpContext(String newHelpContext) {
		myHelpContext = newHelpContext;
	}
	public String $HelpContext() {
		return myHelpContext;
	}
	public void $HelpFile(String newHelpFile) {
		myHelpFile = newHelpFile;
	}
	public String $HelpFile() {
		return myHelpFile;
	}
	public long LastDllError() {
		return 0L;
	}
	public void $Number(int newNumber) {
		myNumber = newNumber;
	}
	public int $Number() {
		return myNumber;
	}
	public void $Source(String newSource) {
		mySource = newSource;
	}
	public String $Source() {
		return mySource;
	}
		
	public void Raise(int number) throws Exception {
		throw Conversion.CVErr(number);
	}
	public void Raise(int number, String Source, String Description, String HelpFile, String HelpContext) throws Exception {
		mySource = Source;
		myDescription = Description;
		myHelpFile = HelpFile;
		myHelpContext = HelpContext;
		myNumber = number;
	}	
}
/*
Throwable
AclNotFoundException, 
ActivationException, 
AlreadyBoundException, 
ApplicationException, 
AWTException, 
BackingStoreException, 
BadAttributeValueExpException, 
BadBinaryOpValueExpException, 
BadLocationException, 
BadStringOperationException, 
BrokenBarrierException, 
CertificateException, 
ClassNotFoundException, 
CloneNotSupportedException, 
DataFormatException, 
DatatypeConfigurationException, 
DestroyFailedException, 
ExecutionException, 
ExpandVetoException, 
FontFormatException, 
GeneralSecurityException, 
GSSException, 
IllegalAccessException, 
IllegalClassFormatException, 
InstantiationException, 
InterruptedException, 
IntrospectionException, 
InvalidApplicationException, 
InvalidMidiDataException, 
InvalidPreferencesFormatException, 
InvalidTargetObjectTypeException, 
InvocationTargetException, 
IOException, 
JMException, 
LastOwnerException, 
LineUnavailableException, 
MidiUnavailableException, 
MimeTypeParseException, 
NamingException, 
NoninvertibleTransformException, 
NoSuchFieldException, 
NoSuchMethodException, 
NotBoundException, 
NotOwnerException, 
ParseException, 
ParserConfigurationException, 
PrinterException, 
PrintException, 
PrivilegedActionException, 
PropertyVetoException, 
RefreshFailedException, ´
RemarshalException, 
RuntimeException, 
SAXException, 
ServerNotActiveException, 
SQLException, 
TimeoutException, 
TooManyListenersException, 
TransformerException, 
UnmodifiableClassException, 
UnsupportedAudioFileException, 
UnsupportedCallbackException, 
UnsupportedFlavorException, 
UnsupportedLookAndFeelException, 
URISyntaxException, 
UserException, 
XAException, 
XMLParseException, 
XPathException 
*/