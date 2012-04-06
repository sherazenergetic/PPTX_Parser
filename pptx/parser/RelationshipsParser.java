package pptx.parser;

import java.io.InputStream;
import java.util.HashMap;
import java.util.zip.ZipEntry;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import pptx.PPTXConstants;
import pptx.context.PPTXContext;
import pptx.exception.BlockerException;


public class RelationshipsParser implements Parser
{
	private ZipEntry relsZipEntry = null;
	private HashMap<String, String> mainRelationShips = null;
	
	public void populateData(ParserInputData parserInputData)  throws BlockerException
	{
		if(parserInputData != null)
		{
			if(parserInputData.getRelsZipEntry() != null)
			{
				this.relsZipEntry = parserInputData.getRelsZipEntry();
			}
		}
	}

	public boolean parse() throws BlockerException
	{
		try
		{				
			if(relsZipEntry == null)
			{
				return false;
			}
			mainRelationShips = new HashMap<String, String>();
			InputStream inputStream = PPTXContext.getInstance().getXMLChunkStream(relsZipEntry);

			Document relsDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			Element root = relsDom.getDocumentElement();
			NodeList children = root.getChildNodes();
			int noOfChildren = children.getLength();
			String key = "", value= "";

			for(int i = 0; i < noOfChildren; i++)
			{	
				Element relation = (Element) children.item(i);
				key = relation.getAttribute(PPTXConstants.PPTX_ATT_RELATIONSHIP_TYPE);
				String relsFileName = relsZipEntry.getName();
				if (!relsFileName.equalsIgnoreCase("_rels/.rels"))
				{
					if (relsFileName.contains("header")
							|| relsFileName.contains("footer"))
					{
						key = "headerFooter"+ key + relation.getAttribute(PPTXConstants.PPTX_ATT_RELATIONSHIP_ID);
					}
					else
					{
						key = relsZipEntry.getName()+ key + relation.getAttribute(PPTXConstants.PPTX_ATT_RELATIONSHIP_ID);
					}
					
				}
				value  = relation.getAttribute(PPTXConstants.PPTX_ATT_RELATIONSHIP_TARGET);
				mainRelationShips.put(key, value);
			}// end of for
		}
		catch(Exception e)
		{
			BlockerException blockerExp = new BlockerException("[RelationshipsParser] Error while parsing source file: " + relsZipEntry.getName() + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		return true;
	}


	public ParserOutputData getOutputData() throws BlockerException
	{
		if(mainRelationShips != null)
		{
			ParserOutputData parserOutputData = new ParserOutputData();
			parserOutputData.setMainRelationShips(mainRelationShips);
			return parserOutputData;
		}
		return null;
	}
}
