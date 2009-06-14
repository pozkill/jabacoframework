package com.jbcdoclet;

import java.io.FileWriter;
import java.io.IOException;

import java.util.HashMap;
import java.util.Vector;
import java.util.Iterator;
import java.util.regex.*;

public class XMLNode {
	private static final String crlf = System.getProperty("line.separator");
	private String _type;
	private static String _processingInstructions = "";
	private static final String _namespace = "http://www.jabaco.org/jabadoc";
	private static String _namespacePrefix = "";
	private HashMap _attributes;
	private Vector _nodes;
	private StringBuffer _text;

	public XMLNode(String type) {
		_type = type;
		_attributes = new HashMap();
		_nodes = new Vector();
		_text = new StringBuffer();
	}
	public void addAttribute(String name, String value) {
		_attributes.put( name, value );
	}
	public String getAttribute(String name) {
		return (String) _attributes.get( name );
	}
	public void addNode( XMLNode node )	{
		_nodes.add( node );
	}
	public void addText( String text ) {
		_text.append( text );
	}
	public void save(String dir, String fileName, boolean includeNamespace)	{
		this.save(dir, fileName, includeNamespace, "");
	}
	public void save(String dir, String fileName, boolean includeNamespace, String outputEncoding) {
		try {
			if(includeNamespace) {
				if(outputEncoding.equals("") == true)
					outputEncoding = "UTF-8";
				_processingInstructions = "<?xml version=\"1.0\" encoding=\"" + outputEncoding + "\" standalone=\"yes\"?>" + crlf;
				_namespacePrefix = "xs";
				this.addAttribute("xmlns:" + _namespacePrefix, _namespace);
				_namespacePrefix = _namespacePrefix + ":";
			} else {
				if(outputEncoding.equals("") == false)
					_processingInstructions = "<?xml version=\"1.0\" encoding=\"" + outputEncoding + "\"?>" + crlf;
			}
			FileWriter out = new FileWriter( dir + fileName );
			out.write( _processingInstructions );
			out.write( this.toString("") );
			out.close();
		} catch( IOException e ) {
			System.out.println( "Could not create '" + dir + fileName + "'" );
		}
	}
	public String toString(String tabs) {
		StringBuffer out = new StringBuffer();
		out.append( tabs + "<" + _namespacePrefix + _type );
		Iterator attrIterator = _attributes.keySet().iterator();
		while( attrIterator.hasNext() ) {
			String key = (String)attrIterator.next();
			out.append( " " + key + "=\"" + encode( (String)_attributes.get( key ) ) + "\"" );
		}

		Iterator nodeIterator = _nodes.iterator();
		
		if( _text.length() <= 0 && !nodeIterator.hasNext() ) {
			out.append( " />" + crlf); 
			return out.toString();
		}
		
		out.append( ">" + crlf);  
		if( _text.length() > 0 ) {
			//Wrapping text in a seperate node allows for good presentation of data with out adding extra data.
			out.append( tabs + "\t<" + _namespacePrefix + "description>" + encode( _text.toString() ) + 
								"</" + _namespacePrefix + "description>" + crlf ); 
		}

		while( nodeIterator.hasNext() )	{
			XMLNode node = (XMLNode)nodeIterator.next();
			out.append( node.toString(tabs + "\t") );
		}

		out.append( tabs + "</" + _namespacePrefix + _type + ">" + crlf  + 
				( "class".equalsIgnoreCase( _type ) ? crlf : "" ));
		
		return out.toString();
	}

	static protected String encode( String in )	{
		Pattern ampPat = Pattern.compile( "&" );
		Pattern ltPat = Pattern.compile( "<" );
		Pattern gtPat = Pattern.compile( ">" );
		Pattern aposPat = Pattern.compile( "\'" );
		Pattern quotPat = Pattern.compile( "\"" );

		String out = new String(in);
		
		/* try { 
			out = new String( in.getBytes(JBCDoclet.getOutputEncoding())) ) ;
		} catch (Exception e) { 
			out = "FAILED!";
		} */

		out = (ampPat.matcher(out)).replaceAll("&amp;");
		out = (ltPat.matcher(out)).replaceAll("&lt;");
		out = (gtPat.matcher(out)).replaceAll("&gt;");
		out = (aposPat.matcher(out)).replaceAll("&apos;");
		out = (quotPat.matcher(out)).replaceAll("&quot;");

		return out;
	}
}
