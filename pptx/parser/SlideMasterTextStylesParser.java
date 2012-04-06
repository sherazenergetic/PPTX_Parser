/**
 * 
 */
package pptx.parser;

import java.io.InputStream;
import java.util.HashMap;
import java.util.zip.ZipEntry;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import pptx.PPTXConstants;
import pptx.context.PPTXContext;
import pptx.exception.BlockerException;
import pptx.pptxObject.PPTXDefaultParagraphStyle;
import pptx.pptxObject.PPTXShape;
import pptx.pptxObject.PPTXSlideMasterTextStyles;

/**
 * @author sheraz.ahmed
 *
 */
public class SlideMasterTextStylesParser implements Parser 
{
	private ZipEntry slideMasterZipEntry = null;
	private PPXMLObject PPTXSlideMasterTextStylesPPXMLObject = null;
	private PPTXContext context = PPTXContext.getInstance();
	private PPXMLObjectFactory ppXMLObjectFactory = context.getFactory();
	

	/**
	 * @see pptx.parser.Parser#populateData(pptx.parser.ParserInputData)
	 */
	public void populateData(ParserInputData parserInputData)
			throws BlockerException 
	{
		if(parserInputData != null)
		{
			if(parserInputData.getSlideMasterZipEntry() != null)
			{
				this.slideMasterZipEntry = parserInputData.getSlideMasterZipEntry();
			}
		}
	}

	/**
	 * @see pptx.parser.Parser#parse()
	 */
	public boolean parse() throws BlockerException 
	{
    	try
		{
    		this.PPTXSlideMasterTextStylesPPXMLObject = ppXMLObjectFactory.createObject(ppXML.PPXML_MAINMASTERS);
    		PPXMLObject mainMaster = ppXMLObjectFactory.createObject(ppXML.PPXML_MAINMASTER);

    		InputStream inputStream = PPTXContext.getInstance().getXMLChunkStream(this.slideMasterZipEntry);
			Document documentDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			
			PPTXSlideMasterTextStyles pptxSlideMasterTextStyles = new PPTXSlideMasterTextStyles();
			Element txStyles = (Element) documentDom.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES).item(0);
			if(txStyles != null && txStyles.getChildNodes().getLength() > 0)
			{
				NodeList txStylesChildren = txStyles.getChildNodes();
				for(int index = 0; index < txStylesChildren.getLength(); index++)
				{
					Node child = txStylesChildren.item(index);
					String childName = child.getNodeName();
					if(childName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES_TITLE_STYLE))
					{
						PPXMLObject masterStyle = parseTitleStyle(child, pptxSlideMasterTextStyles);
						mainMaster.addChild(masterStyle);
					}
					else if(childName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES_BODY_STYLE))
					{
						PPXMLObject masterStyle = parseBodyStyles(child, pptxSlideMasterTextStyles);
						mainMaster.addChild(masterStyle);
					}
					else if(childName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES_OTHER_STYLE))
					{
						PPXMLObject masterStyle = parseOtherStyles(child, pptxSlideMasterTextStyles);
						mainMaster.addChild(masterStyle);
					}
				}// end of for loop
			}
    		this.PPTXSlideMasterTextStylesPPXMLObject.addChild(mainMaster);
    		context.setPPTXSlideMasterTextStyles(pptxSlideMasterTextStyles);

    		Element shapeTree = (Element) documentDom.getElementsByTagName(PPTXConstants.PPTX_SLIDE_SHAPE_TREE).item(0);
    		if(shapeTree != null)
    		{
				NodeList shapes = shapeTree.getElementsByTagName(PPTXConstants.PPTX_SLIDE_SHAPE);
				if(shapes != null && shapes.getLength() > 0)
				{
					HashMap<String, PPTXShape> masterSlideShapes = new HashMap<String, PPTXShape>();
					for(int spIndex = 0; spIndex < shapes.getLength(); spIndex++)
					{
						//p:sp Section 4.4.1.40
						Node shape = shapes.item(spIndex);
						Element shapeElement = (Element) shape;
	
						PPTXShape pptxShape = new PPTXShape();
						pptxShape.populateAttributeValues(shapeElement.getAttributes());
						pptxShape.parseChildren(shapeElement.getChildNodes());
						String type = pptxShape.getPptxNonVisualProperties().getPptxPlaceholderShape().getType();
						masterSlideShapes.put(type, pptxShape);									
					}
					context.setPptxMasterSlideShapes(masterSlideShapes);
				}
    		}
		}
    	catch(Exception e){
    		BlockerException blockerExp = new BlockerException("Error while parsing source file - " + this.slideMasterZipEntry.getName() + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
    	}
    	return true;
	}

	/**
	 * @see pptx.parser.Parser#getOutputData()
	 */
	public ParserOutputData getOutputData() throws BlockerException 
	{
		if(PPTXSlideMasterTextStylesPPXMLObject != null)
		{
			ParserOutputData parserOutputData = new ParserOutputData();
			parserOutputData.setPPTXSlideMasterTextStyles(PPTXSlideMasterTextStylesPPXMLObject);
			return parserOutputData;
		}
		return null;
	}
	
	private PPXMLObject parseTitleStyle(Node titleStyle, PPTXSlideMasterTextStyles pptxSlideMasterTextStyles) throws BlockerException
	{
		PPXMLObject masterStyle = ppXMLObjectFactory.createObject(ppXML.PPXML_MASTERSTYLE);

		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_LEVELS, "1");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_LEVELS + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_TYPE, "Title Style");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_TYPE + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		/*
			- <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
				- <a:spcBef>
				  <a:spcPct val="0" /> 
				  </a:spcBef>

				  <a:buNone />
				   
				- <a:defRPr sz="4400" kern="1200">
					- <a:solidFill>
					  <a:schemeClr val="tx1" /> 
					  </a:solidFill>

					  <a:latin typeface="+mj-lt" /> 
					  <a:ea typeface="+mj-ea" /> 
					  <a:cs typeface="+mj-cs" /> 
				  </a:defRPr>
			  </a:lvl1pPr>
 
		 */
		PPTXDefaultParagraphStyle lvl1pPr = pptxSlideMasterTextStyles.getPptxTitleStyles().getpPTXListTextStyles().getLvl1pPr();
		if(titleStyle != null)
		{
			Element titleStyleElement = (Element) titleStyle;
			if(titleStyleElement.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES_LEVEL1_PPR).item(0) != null)
			{
				Element level1PPR = (Element) titleStyleElement.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_TX_STYLES_LEVEL1_PPR).item(0);
				lvl1pPr.populateAttributeValues(level1PPR.getAttributes());
				lvl1pPr.parseChildren(level1PPR.getChildNodes());
				PPXMLObject ppxmlStyle = lvl1pPr.getPPXMLStyleElement();
				masterStyle.addChild(ppxmlStyle);
			}
		}
		
		return masterStyle;
	}

	private PPXMLObject parseBodyStyles(Node bodyStyle, PPTXSlideMasterTextStyles pptxSlideMasterTextStyles) throws BlockerException
	{
		PPXMLObject masterStyle = ppXMLObjectFactory.createObject(ppXML.PPXML_MASTERSTYLE);
		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_LEVELS, "9");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_LEVELS + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_TYPE, "Body Styles");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_TYPE + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		/*
			- <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
				- <a:spcBef>
  					<a:spcPct val="20000" /> 
  				  </a:spcBef>
  				  <a:buFont typeface="Arial" pitchFamily="34" charset="0" /> 
  				  <a:buChar char="•" /> 
				- <a:defRPr sz="3200" kern="1200">
					- <a:solidFill>
  						<a:schemeClr val="tx1" /> 
  					  </a:solidFill>
  					  <a:latin typeface="+mn-lt" /> 
  					  <a:ea typeface="+mn-ea" /> 
  					  <a:cs typeface="+mn-cs" /> 
  				  </a:defRPr>	
  			  </a:lvl1pPr>
 
		 */
		pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().parseChildren(bodyStyle.getChildNodes());
		
		PPTXDefaultParagraphStyle lvl1pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl1pPr();
		PPXMLObject ppxmlStyle = lvl1pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl2pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl2pPr();
		ppxmlStyle = lvl2pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl3pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl3pPr();
		ppxmlStyle = lvl3pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl4pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl4pPr();
		ppxmlStyle = lvl4pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl5pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl5pPr();
		ppxmlStyle = lvl5pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl6pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl6pPr();
		ppxmlStyle = lvl6pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl7pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl7pPr();
		ppxmlStyle = lvl7pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl8pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl8pPr();
		ppxmlStyle = lvl8pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		PPTXDefaultParagraphStyle lvl9pPr = pptxSlideMasterTextStyles.getPptxBodyStyles().getpPTXListTextStyles().getLvl9pPr();
		ppxmlStyle = lvl9pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
		
		return masterStyle;
	}

	private PPXMLObject parseOtherStyles(Node otherStyle, PPTXSlideMasterTextStyles pptxSlideMasterTextStyles) throws BlockerException
	{
		PPXMLObject masterStyle = ppXMLObjectFactory.createObject(ppXML.PPXML_MASTERSTYLE);
		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_LEVELS, "9");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_LEVELS + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		try
		{
			masterStyle.setProperty(ppXML.PPXML_MASTERSTYLE_TYPE, "Other Styles");
		}catch(Exception e)
		{
    		BlockerException blockerExp = new BlockerException("Error while setting invalid property - " + ppXML.PPXML_MASTERSTYLE_TYPE + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		/*
		- <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
			- <a:spcBef>
					<a:spcPct val="20000" /> 
				  </a:spcBef>
				  <a:buFont typeface="Arial" pitchFamily="34" charset="0" /> 
				  <a:buChar char="•" /> 
			- <a:defRPr sz="3200" kern="1200">
				- <a:solidFill>
						<a:schemeClr val="tx1" /> 
					  </a:solidFill>
					  <a:latin typeface="+mn-lt" /> 
					  <a:ea typeface="+mn-ea" /> 
					  <a:cs typeface="+mn-cs" /> 
				  </a:defRPr>	
			  </a:lvl1pPr>

	     */
		pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().parseChildren(otherStyle.getChildNodes());
		
		PPTXDefaultParagraphStyle lvl1pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl1pPr();
		PPXMLObject ppxmlStyle = lvl1pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl2pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl2pPr();
		ppxmlStyle = lvl2pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
		
		PPTXDefaultParagraphStyle lvl3pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl3pPr();
		ppxmlStyle = lvl3pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl4pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl4pPr();
		ppxmlStyle = lvl4pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl5pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl5pPr();
		ppxmlStyle = lvl5pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl6pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl6pPr();
		ppxmlStyle = lvl6pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl7pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl7pPr();
		ppxmlStyle = lvl7pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl8pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl8pPr();
		ppxmlStyle = lvl8pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);
	
		PPTXDefaultParagraphStyle lvl9pPr = pptxSlideMasterTextStyles.getPptxOtherStyles().getpPTXListTextStyles().getLvl9pPr();
		ppxmlStyle = lvl9pPr.getPPXMLStyleElement();
		masterStyle.addChild(ppxmlStyle);

		return masterStyle;
	}
	

}
