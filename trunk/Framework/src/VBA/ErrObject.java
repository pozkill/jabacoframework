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
	private static PrintStream myErr = null;
	private static int myNumber = 0;
	private static String mySource = null;
	private static String myDescription = null;
	private static String myHelpFile = null;
	private static String myHelpContext = null;
	private static ErrObject myInstance = null;
	
	private ErrObject() {
		//myErr = java.lang.System.err();
	}
	public static void Clear() {
		myNumber = 0;
		mySource = null;
		myDescription = null;
		myHelpFile = null;
		myHelpContext = null;
	}
	public static void $Description(String newDescription) {
		myDescription = newDescription;
	}
	public static String $Description() {
		return myDescription;
	}
	public static void $HelpContext(String newHelpContext) {
		myHelpContext = newHelpContext;
	}
	public static String $HelpContext() {
		return myHelpContext;
	}
	public static void $HelpFile(String newHelpFile) {
		myHelpFile = newHelpFile;
	}
	public static String $HelpFile() {
		return myHelpFile;
	}
	public static long LastDllError() {
		return 0L;
	}
	public static void $Number(int newNumber) {
		myNumber = newNumber;
	}
	public static int $Number() {
		return myNumber;
	}
	public static void $Source(String newSource) {
		mySource = newSource;
	}
	public static String $Source() {
		return mySource;
	}
	public static void Raise(java.lang.Throwable e) throws Throwable {
		throw e;
	}		
	public static void Raise(int number) throws Throwable {
		throw Conversion.ErrNumberToThrowable(number, "");
	}
	public static void Raise(int number, String Source, String Description, String HelpFile, String HelpContext) throws Throwable {
		mySource = Source;
		myDescription = Description;
		myHelpFile = HelpFile;
		myHelpContext = HelpContext;
		myNumber = number;
		Exception e;
		e = new Exception(myDescription);
		throw e;
	}
	public static ErrObject getInstance() {
		if (myInstance == null) myInstance = new ErrObject();
		return myInstance;
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