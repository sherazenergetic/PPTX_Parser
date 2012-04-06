/**
 * 
 */
package pptx.parser;

import java.util.HashMap;

import ppXML.PPXMLFrame;
import ppXML.PPXMLObject;
import ppXML.PPXMLPage;
import ppXML.PPXMLParagraph;
import ppXML.PPXMLShape;
import ppXML.PPXMLShapeGroup;
import pptx.pptxObject.PPTXSlideMasterTextStyles;

/**
 * @author sheraz.ahmed
 *
 */
public class ParserOutputData {
	
	private HashMap<String, String> mainRelationShips = null;
	private PPXMLObject metaTags = null;
	private PPXMLObject PPXMLBody = null; 
	private PPXMLPage page = null;
	private PPXMLShape shape = null;
	private PPXMLFrame ppxmlFrame = null;
	private PPXMLParagraph PPXMLParagraph = null; 
	private PPXMLObject PPTXSlideMasterTextStyles = null;
	private PPXMLShapeGroup shapeGroup = null;
	
	/**
	 * 
	 * @param mainRelationShips
	 */
	public void setMainRelationShips(HashMap<String, String> mainRelationShips) 
	{
		this.mainRelationShips = mainRelationShips;
	}

	/**
	 * 
	 * @return mainRelationShips
	 */
	public HashMap<String, String> getMainRelationShips() 
	{
		return mainRelationShips;
	}
	
	/**
	 * 
	 * @return metaTags
	 */
	public PPXMLObject getMetaTags() 
	{
		return metaTags;
	}

	/**
	 * 
	 * @param metaTags
	 */
	public void setMetaTags(PPXMLObject metaTags) 
	{
		this.metaTags = metaTags;
	}

	/**
	 * 
	 * @return PPXMLBody
	 */
	public PPXMLObject getPPXMLBody() 
	{
		return PPXMLBody;
	}

	/**
	 * 
	 * @param PPXMLBody
	 */
	public void setPPXMLBody(PPXMLObject PPXMLBody) 
	{
		this.PPXMLBody = PPXMLBody;
	}

	/**
	 * 
	 * @return page instance of PPXMLPage
	 */
	public PPXMLPage getPage() 
	{
		return page;
	}

	/**
	 * 
	 * @param page instance of PPXMLPage
	 */
	public void setPage(PPXMLPage page) 
	{
		this.page = page;
	}

	/**
	 * 
	 * @return shape instance of PPXMLShape
	 */
	public PPXMLShape getShape() 
	{
		return shape;
	}

	/**
	 * 
	 * @param shape instance of PPXMLShape
	 */
	public void setShape(PPXMLShape shape) 
	{
		this.shape = shape;
	}

	/**
	 * 
	 * @return ppxmlFrame
	 */
	public PPXMLFrame getPPPXMLFrame() 
	{
		return ppxmlFrame;
	}

	/**
	 * 
	 * @param ppxmlFrame
	 */
	public void setPPPXMLFrame(PPXMLFrame ppxmlFrame) 
	{
		this.ppxmlFrame = ppxmlFrame;
	}

	/**
	 * 
	 * @return PPXMLParagraph
	 */
	public PPXMLParagraph getPPXMLParagraph() {
		return PPXMLParagraph;
	}

	/**
	 * 
	 * @param PPXMLParagraph
	 */
	public void setPPXMLParagraph(PPXMLParagraph PPXMLParagraph) {
		this.PPXMLParagraph = PPXMLParagraph;
	}

	/**
	 * @return the pPTXSlideMasterTextStyles
	 */
	public PPXMLObject getPPTXSlideMasterTextStyles() {
		return PPTXSlideMasterTextStyles;
	}

	/**
	 * @param pPTXSlideMasterTextStyles the pPTXSlideMasterTextStyles to set
	 */
	public void setPPTXSlideMasterTextStyles(
			PPXMLObject pPTXSlideMasterTextStyles) {
		this.PPTXSlideMasterTextStyles = pPTXSlideMasterTextStyles;
	}

	/**
	 * @return Returns the shapeGroup.
	 */
	public PPXMLShapeGroup getShapeGroup() {
		return shapeGroup;
	}

	/**
	 * @param shapeGroup The shapeGroup to set.
	 */
	public void setShapeGroup(PPXMLShapeGroup shapeGroup) {
		this.shapeGroup = shapeGroup;
	}

	
}
