package pptx.parser;

import java.io.InputStream;
import java.util.zip.ZipEntry;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import pptx.context.PPTXContext;
import pptx.exception.BlockerException;
import ppXML.PPXMLMeta;
import ppXML.PPXMLObject;
import ppXML.PPXMLObjectFactory;

public class AppExtendedPropertiesParser implements Parser {

	private ZipEntry appZipEntry = null;
	private PPXMLObject metaTags = null;
	
	public void populateData(ParserInputData parserInputData) throws BlockerException 
	{
		if(parserInputData != null)
		{
			if(parserInputData.getAppZipEntry() != null)
			{
				this.appZipEntry = parserInputData.getAppZipEntry();
			}
		}
	}

	public boolean parse() throws BlockerException
    {
    	PPTXContext context = PPTXContext.getInstance();
    	PPXMLObjectFactory factory = context.getFactory();
    	this.metaTags = factory.createObject(ppXML.PPXML_DOCUMENT_METATAGS);
    	try
		{
	    	InputStream inputStream = PPTXContext.getInstance().getXMLChunkStream(this.appZipEntry);
			Document documentDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			Element appProperties = documentDom.getDocumentElement();			
			NodeList children = appProperties.getChildNodes();
			int noOfChildren = children.getLength();
			for(int i = 0; i < noOfChildren; i++)
			{
				PPXMLMeta metaElement = (PPXMLMeta)factory.createObject(ppXML.PPXML_META);				
				if (children.item(i).getLocalName() != null && XercesUtils.getTextContent(children.item(i)) != null ){
				metaElement.setProperty(ppXML.PPXML_ATT_NAME, children.item(i).getLocalName());
				metaElement.setProperty(ppXML.PPXML_ATT_VALUE, XercesUtils.getTextContent(children.item(i)));				
				metaTags.addChild(metaElement);	
			}
		}
		}
    	catch(Exception e){
    		BlockerException blockerExp = new BlockerException("Error while parsing source file properties: " + appZipEntry.getName() + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
    	}
    	
    	return true;
    }


	public ParserOutputData getOutputData() throws BlockerException 
	{
		if(metaTags != null)
		{
			ParserOutputData parserOutputData = new ParserOutputData();
			parserOutputData.setMetaTags(metaTags);
			return parserOutputData;
		}
		return null;
	}
}
