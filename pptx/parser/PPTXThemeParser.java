/**
 * 
 */
package pptx.parser;

import java.io.InputStream;
import java.util.zip.ZipEntry;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


/**
 * @author sheraz.ahmed
 *
 */
public class PPTXThemeParser implements Parser 
{
	private ZipEntry masterSlideThemeZipEntry = null;
	private PPTXContext context = PPTXContext.getInstance();
	/**
	 * @see pptx.parser.Parser#populateData(pptx.parser.ParserInputData)
	 * @param parserInputData
	 */
	public void populateData(ParserInputData parserInputData)
			throws BlockerException 
	{
		if(parserInputData != null)
		{
			if(parserInputData.getMasterSlideThemeZipEntry() != null)
			{
				this.masterSlideThemeZipEntry = parserInputData.getMasterSlideThemeZipEntry();
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
	    	InputStream inputStream = PPTXContext.getInstance().getXMLChunkStream(this.masterSlideThemeZipEntry);
			Document documentDom = XYZDOMUtilities.parseXMLFromInputStream(inputStream);
			Element themeElementsElement = (Element)documentDom.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_THEME_ELEMENTS).item(0);
			if(themeElementsElement != null)
			{
				NodeList children = themeElementsElement.getChildNodes();
				int noOfChildren = children.getLength();
				for(int index = 0; index < noOfChildren; index++)
				{
					Node child = children.item(index);
					String nodeName = child.getNodeName();
					if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDEMASTER_THEME_COLOR_SCHEME))
					{
						context.loadPPTXColors(child.getChildNodes());
					}
					else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDEMASTER_THEME_FONT_SCHEME))
					{
						Element majorFontElement = (Element)documentDom.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_THEME_COLOR_SCHEME_MAJORFONT).item(0);
						if(majorFontElement != null)
						{
							context.loadPPTXMajorFonts(majorFontElement.getChildNodes());
						}
						Element minorFontElement = (Element)documentDom.getElementsByTagName(PPTXConstants.PPTX_SLIDEMASTER_THEME_COLOR_SCHEME_MINORFONT).item(0);
						if(minorFontElement != null)
						{
							context.loadPPTXMinorFonts(minorFontElement.getChildNodes());
						}
					}
				}// end of for loop
			}
		}
    	catch(Exception e){
    		BlockerException blockerExp = new BlockerException("Error while parsing source file - " + this.masterSlideThemeZipEntry.getName() + " - " + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
    	}

		return true;
	}

	/**
	 * @see pptx.parser.Parser#getOutputData()
	 * @return ParserOutputData
	 */
	public ParserOutputData getOutputData() throws BlockerException 
	{
		// TODO Auto-generated method stub
		return null;
	}

}
