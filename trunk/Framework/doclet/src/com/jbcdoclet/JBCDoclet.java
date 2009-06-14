package com.jbcdoclet;

import com.sun.javadoc.*;
import java.util.Date;

public class JBCDoclet {

	private static String programversion = "2.1.0";
	private static String xsdversion = "1.0.0";
	private static boolean multipleFiles = false;
	private static String outputDirectory = "./";
	private static boolean includeNamespace = false;
	private static String outputEncoding = "UTF-8";
	private static String filenameBase = "";

	public static String getOutputEncoding() {
		return outputEncoding;
	}
	
	public static boolean start(RootDoc root) {
		getOptions(root);
		XMLNode[] nodes = buildXmlFromDoc(root);
		save(nodes);
		return true;
	}

	public static int optionLength( String option ) {
		if ( option.compareToIgnoreCase( "-multiple" ) == 0 )			return 1;
		if ( option.compareToIgnoreCase( "-includeNamespace" ) == 0 )	return 1;
		if ( option.compareToIgnoreCase( "-d" ) == 0 )					return 2;
		if ( option.compareToIgnoreCase( "-outputEncoding" ) == 0 )		return 2;
		if ( option.compareToIgnoreCase( "-filename" ) == 0 )			return 2;
		return 0;
	}

	public static boolean validOptions(String options[][], DocErrorReporter reporter )	{	
		return true;
	}
	
	private static void getOptions(RootDoc root) {
		String[][] options = root.options();
		for ( int opt = 0; opt < options.length; opt++ ) {
			if( options[opt][0].compareToIgnoreCase( "-d" ) == 0 ) {
				outputDirectory = options[opt][1];
				String fs = System.getProperty("file.separator");
				if(outputDirectory.endsWith(fs) == false)
					outputDirectory += fs;
				continue;
			}
			if( options[opt][0].compareToIgnoreCase( "-multiple" ) == 0 ) {
				multipleFiles = true;
				continue;
			}
			if( options[opt][0].compareToIgnoreCase( "-includeNamespace" ) == 0 ) {
				includeNamespace = true;
				continue;
			}
			if( options[opt][0].compareToIgnoreCase( "-outputEncoding" ) == 0 )	{
				outputEncoding = options[opt][1];
				continue;
			}
			if( options[opt][0].compareToIgnoreCase( "-filename" ) == 0 ) {
				filenameBase = options[opt][1];
				continue;
			}
		}
	    System.out.println("jbcdoclet V" + programversion);
		System.out.println( "Using output directory '" + outputDirectory + "'.");
		System.out.println( "output encoding '" + outputEncoding + "'.");
		System.out.println("Saving as " + (multipleFiles ? "multiple files." : "a single file.") );

		if(filenameBase.equals("") == false)
			System.out.println("filename " + (multipleFiles ? " base " : "") + ": '" + filenameBase + "'");
	}
	

	private static void save(XMLNode[] nodes) {
		XMLNode rootXml = new XMLNode("jabadoc");
		XMLNode adminNode = new XMLNode("admin");
		adminNode.addAttribute("version", programversion);
		adminNode.addAttribute("xsdversion", xsdversion);
		adminNode.addAttribute("creation", new Date().toString());
		rootXml.addNode(adminNode);
		if(multipleFiles) {
			for(int index=0; index<nodes.length; index++) {
				String fileName = nodes[index].getAttribute("fulltype") + ".xml";
				rootXml.addNode( nodes[index] );
				if(filenameBase.equals("") == false)
					fileName = filenameBase + fileName;
				rootXml.save(outputDirectory, fileName, includeNamespace, outputEncoding);
			}
		} else {
			for(int index=0; index<nodes.length; index++)
				rootXml.addNode(nodes[index]);
				String fileName = "DocReport.xml";
				if(filenameBase.equals("") == false) fileName = filenameBase;
				rootXml.save(outputDirectory, fileName, includeNamespace, outputEncoding);
		}
	}
	
	private static XMLNode[] buildXmlFromDoc(RootDoc root) {
		ClassDoc[] classes = root.classes();
		XMLNode[] retval = new XMLNode[classes.length];
		for ( int index = 0; index < classes.length; index++ ) {
			retval[index] = transformClass( classes[index], null );
		}
		return retval;
	}
	
	private static void transformComment(Doc doc, XMLNode node) {
		XMLNode commentNode = new XMLNode( "comment" );
		boolean addNode = false;
		if ( doc.commentText() != null && doc.commentText().length() > 0 ) {
			commentNode.addText( doc.commentText() );
			addNode = true;
		}
		Tag[] tags = doc.tags();
		for( int tag = 0; tag < tags.length; tag++ ) {		
			XMLNode paramNode = new XMLNode( "attribute" );
			paramNode.addAttribute( "name", tags[tag].name() );
			paramNode.addText( tags[tag].text() );
			commentNode.addNode( paramNode );
			addNode = true;
		}
		if ( addNode ) node.addNode( commentNode );
	}

	private static void transformFields(FieldDoc[] fields, XMLNode node) {
		if ( fields.length < 1 ) return;
		XMLNode fieldsNode = new XMLNode( "fields"); 
		for( int index = 0; index < fields.length; index++ ) {
			XMLNode fieldNode = new XMLNode( "field");
			fieldNode.addAttribute( "name", fields[index].name() );
			fieldNode.addAttribute( "type", fields[index].type().typeName() );
			fieldNode.addAttribute( "fulltype", fields[index].type().toString() );
			
			if ( fields[index].constantValue() != null && fields[index].constantValue().toString().length() > 0 )
				fieldNode.addAttribute( "const", fields[index].constantValue().toString() );

			if ( fields[index].constantValueExpression() != null && fields[index].constantValueExpression().length() > 0 )
				fieldNode.addAttribute( "constexpr", fields[index].constantValueExpression() );

			setVisibility(fields[index], fieldNode);
			
			if ( fields[index].isStatic() )
				fieldNode.addAttribute( "static", "true" );

			if ( fields[index].isFinal() )
				fieldNode.addAttribute( "final", "true" );

			if ( fields[index].isTransient() )
				fieldNode.addAttribute( "transient", "true" );

			if ( fields[index].isVolatile() )
				fieldNode.addAttribute( "volatile", "true" );

			transformComment( fields[index], fieldNode );
			fieldsNode.addNode( fieldNode ); 
		}
		node.addNode( fieldsNode );
	}
	
	private static void setVisibility(ProgramElementDoc member, XMLNode node)	{
		if ( member.isPrivate() )
			node.addAttribute( "visibility", "private" );
		
		else if ( member.isProtected() )
			node.addAttribute( "visibility", "protected" );
		
		else if ( member.isPublic() )
			node.addAttribute( "visibility", "public" );
		
		else if ( member.isPackagePrivate() )
			node.addAttribute( "visibility", "package-private" );
	}

	private static void populateMethodNode(ExecutableMemberDoc method, XMLNode node) {
		transformComment( method, node );
		node.addAttribute( "name", method.name() );
		setVisibility(method, node);
		if ( method.isStatic() )
			node.addAttribute( "static", "true" );
		if ( method.isInterface() )
			node.addAttribute( "interface", "true" );
		if ( method.isFinal() )
			node.addAttribute( "final", "true" );
		if ( method instanceof MethodDoc )
			if ( ((MethodDoc) method).isAbstract() )
				node.addAttribute( "abstract", "true" );
		if ( method.isSynchronized() )
			node.addAttribute( "synchronized", "true" );
		if ( method.isSynthetic() )
			node.addAttribute( "synthetic", "true" );
		Parameter[] params = method.parameters();
		if ( params.length > 0 ) {
			ParamTag[] paramTags = method.paramTags();
			XMLNode paramsNode = new XMLNode( "params" );
			for( int param = 0; param < params.length; param++ ) {
				XMLNode paramNode = new XMLNode( "param" );
				paramNode.addAttribute( "name", params[param].name() );
				paramNode.addAttribute( "type", params[param].type().typeName() );
				paramNode.addAttribute( "fulltype", params[param].type().toString() );
				for( int paramTag = 0; paramTag < paramTags.length; paramTag++ ) {
					if( paramTags[ paramTag ].parameterName().compareToIgnoreCase( params[param].name() ) == 0 ) {
						paramNode.addAttribute( "comment", paramTags[ paramTag ].parameterComment() );
					}
				}
				paramsNode.addNode( paramNode );
			}
			node.addNode( paramsNode );
		}
		ClassDoc[] exceptions = method.thrownExceptions();
		ThrowsTag[] exceptionTags = method.throwsTags();
		if ( exceptions != null && exceptions.length > 0 ) {
			XMLNode exceptionsNode = new XMLNode( "exceptions" );
			for( int except = 0; except < exceptions.length; except++ )	{
				XMLNode exceptNode = new XMLNode( "exception" );
				exceptNode.addAttribute( "type", exceptions[except].typeName() );
				exceptNode.addAttribute( "fulltype", exceptions[except].qualifiedTypeName() );
		        for(int exceptionTag = 0; exceptionTag < exceptionTags.length; exceptionTag++) {
					if(exceptionTags[exceptionTag].exceptionName().compareToIgnoreCase(exceptions[except].typeName()) == 0)
		            exceptNode.addAttribute("comment", exceptionTags[exceptionTag].exceptionComment());
				}
					exceptionsNode.addNode( exceptNode );
			}
			node.addNode( exceptionsNode );
		}
	}
	
	private static void transformMethods(MethodDoc[] methods, ConstructorDoc[] constructors, XMLNode node) {
		if ( methods.length < 1 && constructors.length < 1 ) return;
		XMLNode methodsNode = new XMLNode( "methods"); 
		for( int index = 0; index < constructors.length; index++ ) {
			XMLNode constNode = new XMLNode( "constructor");
			populateMethodNode( constructors[index], constNode );
			methodsNode.addNode( constNode ); 
		}
		for( int index = 0; index < methods.length; index++ ) {
			XMLNode methodNode = new XMLNode( "method");
			populateMethodNode( methods[index], methodNode );
			methodNode.addAttribute( "type", methods[index].returnType().typeName() );
			methodNode.addAttribute( "fulltype", methods[index].returnType().toString() );
			Tag[] returnTags = methods[ index ].tags( "@return" );
			if ( returnTags.length > 0 ) {
				methodNode.addAttribute( "returncomment", returnTags[ 0 ].text() );
			}
			methodsNode.addNode( methodNode ); 
		}
		node.addNode( methodsNode );
	}

	private static XMLNode transformClass(ClassDoc classDoc, XMLNode root) {
		XMLNode classNode = new XMLNode( "class" );  //"class" needs a prefix for output to work with JAXB.
		
		// Handle basic class attributes
		setVisibility(classDoc, classNode);
		classNode.addAttribute( "type", classDoc.name() );
		classNode.addAttribute( "fulltype", classDoc.qualifiedName() );
		classNode.addAttribute( "package", classDoc.containingPackage().name() );
		
		ClassDoc[] extendClasses = classDoc.interfaces();
		if(extendClasses.length > 0) {
			XMLNode implement = new XMLNode( "implements" );
			for( int extendIndex = 0; extendIndex < extendClasses.length; extendIndex++ ) {
				XMLNode interfce = new XMLNode( "interface" );
				interfce.addAttribute( "type", extendClasses[extendIndex].name() );
				interfce.addAttribute( "fulltype", extendClasses[extendIndex].qualifiedName() );
				implement.addNode( interfce );
			}
	       classNode.addNode(implement);
		}

		if ( classDoc.superclass() != null ) {
			classNode.addAttribute( "superclass", classDoc.superclass().name() );
			classNode.addAttribute( "superclassfulltype", classDoc.superclass().qualifiedName() );
		}

		if ( classDoc.isInterface() )
			classNode.addAttribute( "interface", "true" );
		
		if ( classDoc.isFinal() )
			classNode.addAttribute( "final", "true" );
			
		if ( classDoc.isAbstract() )
			classNode.addAttribute( "abstract", "true" );
		
		if ( classDoc.isSerializable() )
			classNode.addAttribute( "serializable", "true" );
		
		transformComment( classDoc, classNode );
		transformFields( classDoc.fields(), classNode );
		transformMethods( classDoc.methods(), classDoc.constructors(), classNode );
		ClassDoc[] innerClasses = classDoc.innerClasses();
		for( int classIndex = 0; classIndex < innerClasses.length; classIndex++ )
			classNode.addNode ( transformClass( innerClasses[classIndex], classNode ) );

		return classNode;
	}
}
