package VBA;

public class Constants extends VBEnumClass {
    /** Represents a backspace character for print and display functions. */
	public static final String vbBack = new String( new char[]{8} );
	/** Represents a carriage-return character for print and display functions. */
	public static final String vbCr = new String( new char[]{13} );
	/** Represents a carriage-return character combined with a linefeed character for print and display functions. */
	public static final String vbCrLf = new String( new char[]{13,10});
	/** Represents a form-feed character for print functions. */
	public static final String vbFormFeed = new String( new char[]{12} );
	/** Represents a form-feed character for print functions. */
	public static final String vbLf = new String( new char[]{10} );
	/** Represents a newline character for print and display functions. */
	public static final String vbNewLine = vbCrLf;
	/** Represents a newline character for print and display functions. */
	public static final String vbNullChar = new String( new char[]{0} );;
	/** Represents a zero-length string for print and display functions, and for calling external procedures. */
	public static final String vbNullString = new String();
	/** Represents a tab character for print and display functions. */
	public static final String vbTab = new String( new char[]{9} );
	/** Represents a carriage-return character for print functions. */
	public static final String vbVerticalTab = new String( new char[]{11} );
}
