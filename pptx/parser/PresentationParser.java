/**
 * 
 */
package pptx.parser;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipEntry;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


import pptx.PPTXConstants;
import pptx.context.PPTXContext;
import pptx.exception.BlockerException;
import pptx.pptxObject.PPTXShape;

/**
 * @author sheraz.ahmed
 *
 */
public class PresentationParser implements Parser {
	
	private ZipEntry presentationEntry = null;
	private HashMap<String, String> presentationRelationships = null;
	private HashMap<String, String> slidesRelsAllReletionships = new HashMap<String, String>();
	private PPTXContext context = PPTXContext.getInstance();
	private PPXMLObjectFactory ppXMLObjectFactory = context.getFactory();
	private PPXMLObject ppxmlBody = null;
	private PPTXParserFactory parserFactory = PPTXParserFactory.getInstance();
	private Document slideRelationDom = null;

	/** 
	 * @see pptx.parser.Parser#populateData(pptx.parser.ParserInputData)
	 */
	public void populateData(ParserInputData parserInputData)
			throws BlockerException 
	{
		this.presentationEntry = parserInputData.getPresentationEntry();
		this.presentationRelationships = parserInputData.getPresentationRelationships();
	}

	/**
	 * @see pptx.parser.Parser#parse()
	 */
	public boolean parse() throws BlockerException 
	{
		try
		{
			ppxmlBody = ppXMLObjectFactory.createObject(ppXML.PPXML_BODY);
			HashMap<String, ZipEntry> pptxEntries = context.getPPTXEntries();
			List<String> slideEntries = getAllSlideEntries();
			
			context.setMarker_LastSlideNo(slideEntries.size());
			for(String slideEntry:slideEntries)
			{
				slideEntry = "ppt/" + slideEntry;
				ZipEntry slideZipEntry = pptxEntries.get(slideEntry);
				
				String slideRelationEntry = slideEntry.substring(0, slideEntry.lastIndexOf("/"));
				slideRelationEntry += "/_rels/" + slideEntry.substring(slideEntry.lastIndexOf("/") + 1, slideEntry.length()) + ".rels";
				
				String slideRelationShipTarget = getSlideRelationshipTargetEntry(slideRelationEntry);

				HashMap<String, InputStream> ImagesStreamsMap = getImageMap(slideRelationEntry, pptxEntries);
				
				if(slideRelationShipTarget != null && !(slideRelationShipTarget.equals("")))
				{
					slideRelationShipTarget = slideRelationShipTarget.substring(slideRelationShipTarget.indexOf("/"));
					slideRelationShipTarget = "ppt" + slideRelationShipTarget;
					ZipEntry slideLayoutZipEntry = pptxEntries.get(slideRelationShipTarget);
					PPXMLObject ppxmlPage = parseSlide(slideZipEntry, this.slidesRelsAllReletionships, slideLayoutZipEntry, ImagesStreamsMap);
					ppxmlBody.addChild(ppxmlPage);
				}
				else
				{
					PPXMLObject ppxmlPage = parseSlide(slideZipEntry, this.slidesRelsAllReletionships, null, ImagesStreamsMap);
					ppxmlBody.addChild(ppxmlPage);
				}


				
				
				
				
			}
			
//			InputStream inputStream = PPTXContext.getInstance().getXMLChunkStream(presentationEntry);
//			Document presentationDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
//			
//			Element presentationDocElement = presentationDom.getDocumentElement();
//			
//			Element slideSize = (Element) presentationDocElement.getElementsByTagName(PPTXConstants.PPTX_PRESENTATION_SLIDE_SIZE).item(0);
//			
//			PPXMLObject ppxmlPresentation =  ppXMLObjectFactory.createObject(ppXML.PPXML_PRESENTATION);
//			if(slideSize != null)
//			{
//				ppxmlPresentation.setProperty(ppXML.PPXML_PRESENTATION_CX, slideSize.getAttribute(ppXML.PPXML_PRESENTATION_CX));
//				ppxmlPresentation.setProperty(ppXML.PPXML_PRESENTATION_CY, slideSize.getAttribute(ppXML.PPXML_PRESENTATION_CY));
//				ppxmlPresentation.setProperty(ppXML.PPXML_PRESENTATION_TYPE, slideSize.getAttribute(ppXML.PPXML_PRESENTATION_TYPE));
//			}
//			ppxmlBody.addChild(ppxmlPresentation);
			
		}
		catch(Exception e)
		{
			BlockerException blockerExp = new BlockerException("PPTX - Error while parsing source file : " + presentationEntry.getName() + " - " + e.getMessage());
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
		if(ppxmlBody != null)
		{
			ParserOutputData parserOutputData = new ParserOutputData();
			parserOutputData.setPPXMLBody(ppxmlBody);
			return parserOutputData;
		}
		return null;
	}
	
	private List<String> getAllSlideEntries()
	{
		ArrayList<String> slideEntries = new ArrayList<String>();
		Iterator<String> iterator = presentationRelationships.keySet().iterator();

		while(iterator.hasNext())
		{
			String key = iterator.next();
			String keyValue = key.substring(key.indexOf("rId")+3);
			if(key.contains(PPTXConstants.PPTX_RELATIONSHIPS_SLIDE) && !key.contains(PPTXConstants.PPTX_RELATIONSHIPS_SLIDEMASTER))
			{
				slideEntries.add(keyValue);
			}
		}
		int slideEntryArray[] = new int[slideEntries.size()];
//		slideEntryArray = slideEntries.toArray(slideEntryArray);
		//String sortedSlideEntryArray [] = sortRelionshipsBySlideNumber (slideEntryArray);
		for(int lIndex = 0; lIndex < slideEntries.size(); lIndex++)
		{
			slideEntryArray[lIndex] = Integer.parseInt(slideEntries.get(lIndex));
		}
		Arrays.sort(slideEntryArray);
		slideEntries = new ArrayList<String>();
		for(int i = 0; i < slideEntryArray.length; i++)
		{
			String slideEntry = "" + slideEntryArray[i]; 
			String actualEntry = reterieveActualEntry(slideEntry);
			if(presentationRelationships.containsKey(actualEntry))
			{
				slideEntries.add(presentationRelationships.get(actualEntry));
			}
		}
		return slideEntries;
	}
	
	
//	private String[] sortRelionshipsBySlideNumber (String unsortedSlidesByName[])
//	{
//		
//		ArrayList<String> sortedSlideEntries = new ArrayList<String>();
//		sortedSlideEntries.add(unsortedSlidesByName[0]);
//		String sortedSlideEntryArray[]= new String[unsortedSlidesByName.length];
//		//sortedSlideEntryArray[0] = unsortedSlidesByName[0];
//		
//		for(int i = 0; i < unsortedSlidesByName.length; i++)
//		{
//			String unsortedSlideEntry = unsortedSlidesByName[i]; 
//			
//			for(int j = 0; j < sortedSlideEntries.size(); j++)
//			{
//				String sortedSlideEntry = sortedSlideEntries.get(j); 
//				
//				if(Double.parseDouble(unsortedSlideEntry.substring(3)) > Double.parseDouble(sortedSlideEntry.substring(3)))
//				{
//					sortedSlideEntries.add(unsortedSlideEntry);
//				}
//				else
//				{
//					sortedSlideEntries.add(i-1,unsortedSlideEntry);
//				}
//			}
//			
//			
//			
//		}
//		
		
//		
//		return sortedSlideEntryArray;
//	}
//	
	
	private String reterieveActualEntry(String key)
	{
		String actualKey = "";
		Iterator<String> iterator = presentationRelationships.keySet().iterator();

		while(iterator.hasNext())
		{
			String mapKey = iterator.next();
			if(mapKey.contains(PPTXConstants.PPTX_RELATIONSHIPS_SLIDE) && !mapKey.contains(PPTXConstants.PPTX_RELATIONSHIPS_SLIDEMASTER))
			{
				if(mapKey.endsWith("rId" + key))
				{
					actualKey = mapKey;
					break;
				}
			}
			
		}
		return actualKey;
	}
	
	private PPXMLPage parseSlide(ZipEntry slide, HashMap<String, String> slideRels, ZipEntry slideLayoutZipEntry, HashMap<String, InputStream> imagesStreamsMap) throws BlockerException
	{
		PPXMLPage page = null;

		Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_SLIDEPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setSlide(slide);
    	parserInputData.setSlideRelationships(slideRels);
    	parserInputData.setSlideLayout(slideLayoutZipEntry);
    	parserInputData.setImagesStreamsMap(imagesStreamsMap);
    	parser.populateData(parserInputData);
    	

    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		page = parser.getOutputData().getPage();
    	}

		return page;
	}
	
	private String getSlideRelationshipTargetEntry(String slideRelationEntry)
	{
		String slideRelationShipTarget = "";

		ZipEntry slideRelationZipEntry = context.getPPTXEntries().get(slideRelationEntry);				
		if(slideRelationZipEntry != null)
		{
			InputStream inputStream = context.getXMLChunkStream(slideRelationZipEntry);
			this.slideRelationDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			if(this.slideRelationDom != null)
			{
				NodeList slideRelationShips = this.slideRelationDom.getElementsByTagName(PPTXConstants.PPTX_SLIDE_RELATIONSHIP);
				for(int index = 0; index < slideRelationShips.getLength(); index++)
				{
					Node slideRelationShip = slideRelationShips.item(index);
					if(slideRelationShip.getAttributes() != null && slideRelationShip.getAttributes().getLength() > 0)
					{
						NamedNodeMap namedNodeMap = slideRelationShip.getAttributes();
						if(namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TYPE) != null)
						{
							String relationShipType = namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TYPE).getNodeValue();
							this.slidesRelsAllReletionships.put(namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_ID).getNodeValue(),namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TARGET).getNodeValue());
							if(relationShipType != null && relationShipType.equalsIgnoreCase(PPTXConstants.PPTX_RELATIONSHIPS_SLIDE_LAYOUT))
							{
								slideRelationShipTarget = namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TARGET).getNodeValue();
//								break;
							}
						}
					}
				}
			}
		}
		return slideRelationShipTarget;
	}

	private HashMap<String, InputStream> getImageMap(String slideRelationEntry, HashMap<String, ZipEntry> pptxEntries)
	{
//		InputStream inputStream = context.getXMLChunkStream(entry);
		HashMap<String, InputStream> imageStreamsMap = new HashMap<String, InputStream>();  

		ZipEntry slideRelationZipEntry = context.getPPTXEntries().get(slideRelationEntry);				
		if(slideRelationZipEntry != null)
		{
//			InputStream inputStream = context.getXMLChunkStream(slideRelationZipEntry);
//			Document slideRelationDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			if(this.slideRelationDom != null)
			{
				NodeList slideRelationShips = this.slideRelationDom.getElementsByTagName(PPTXConstants.PPTX_SLIDE_RELATIONSHIP);
				for(int index = 0; index < slideRelationShips.getLength(); index++)
				{
					Node slideRelationShip = slideRelationShips.item(index);
					if(slideRelationShip.getAttributes() != null && slideRelationShip.getAttributes().getLength() > 0)
					{
						NamedNodeMap namedNodeMap = slideRelationShip.getAttributes();
						if(namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TYPE) != null)
						{
							String relationShipType = namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TYPE).getNodeValue();
							if(relationShipType != null && relationShipType.equalsIgnoreCase(PPTXConstants.PPTX_RELATIONSHIPS_IMAGE))
							{
								String id = namedNodeMap.getNamedItem("Id").getNodeValue();
								String slideImageTarget = namedNodeMap.getNamedItem(PPTXConstants.PPTX_ATT_RELATIONSHIP_TARGET).getNodeValue();
								slideImageTarget = slideImageTarget.substring(slideImageTarget.indexOf("/"));
								slideImageTarget = "ppt" + slideImageTarget;
								ZipEntry slideImageZipEntry = pptxEntries.get(slideImageTarget);
								InputStream imageInputStream = context.getXMLChunkStream(slideImageZipEntry);
								String key = id + slideImageTarget;
								imageStreamsMap.put(key, imageInputStream);
							}
						}
					}
				}// end of for slideRelationShipEntry
			}
		}
		return imageStreamsMap;
	}

}
