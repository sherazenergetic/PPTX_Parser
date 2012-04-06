
package pptx.pptxObject;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import pptx.PPTXUtils;
import pptx.context.PPTXContext;


/**
 * @author Sheraz Ahmed
 *
 */
public class PPTXTableCellProperties {
	
	private PPTXContext context = PPTXContext.getInstance();
	
	// tcPr (Table Cell Properties) 5.1.6.15
	private String anchor = "";			//Defines the alignment of the text vertically within the cell.
	private boolean anchorCtr = false;  //When this attribute is on, 1 or true, it modifies the anchor attribute
	private String horzOverflow = "";   //Specifies the clipping behavior of the cell.
	private String marB = "";    //Specifies the bottom margin of the cell      
	private String marL = "";	 //Specifies the left margin of the cell
	private String marR = "";    //Specifies the right margin of the cell
	private String marT = "";    //Specifies the top margin of the cell
	private String vert = "";	 //Defines the text direction within the cell
	
	// Children
	//::::1- blipFill (Picture Fill) 5.1.10.14
	private PPTXBlipFill pptxBlipFill = new PPTXBlipFill();
	
	//::::2- cell3D (Cell 3-D) 5.1.6.1
	
	//::::3- extLst (Extension List) 5.1.2.1.15
	private PPTXExtensionList pptxExtensionList = new PPTXExtensionList();
	
	//::::4- gradFill (Gradient Fill) 5.1.10.33
	private PPTXGradientFill pptxGradientFill = new PPTXGradientFill();
	
	//::::5- grpFill (Group Fill) 5.1.10.35
	private boolean groupFill = false;//This element specifies a group fill. When specified, this 
	//setting indicates that the parent element is part of a group and should inherit the 
	//fill properties of the group. 
	
	//::::6- lnB (Bottom Border Line Properties) 5.1.6.3
	private PPTXOutline pptxBottomBorderLineProperties = new PPTXOutline();
	
	//::::7- lnBlToTr (Bottom-Left to Top-Right Border Line Properties) 5.1.6.4
	private PPTXOutline pptxBottomLToTopRightBorderProperties = new PPTXOutline();
	
	//::::8- lnL (Left Border Line Properties) 5.1.6.5
	private PPTXOutline pptxLeftBorderLineProperties = new PPTXOutline();
	
	//::::9- lnT (Top Border Line Properties) 5.1.6.7
	private PPTXOutline pptxTopBorderLineProperties = new PPTXOutline();
	
	//::::10- lnTlToBr (Top-Left to Bottom-Right Border Line Properties) 5.1.6.8
	private PPTXOutline pptxTopLToBottomRightBorderProperties = new PPTXOutline();
	
	//::::11- noFill (No Fill) 5.1.10.44
	private boolean noFill = false;//This element specifies that no fill will be applied to the parent element.
	
	//::::12- pattFill (Pattern Fill) 5.1.10.47
	private String presetPattern = ""; //Specifies one of a set of preset patterns to fill the object.
	//children
	// bgClr (Background color) 5.1.10.10
	private PPTXColor bgColor = new PPTXColor();
	// fgClr (Foreground color) 5.1.10.27
	private PPTXColor fgColor = new PPTXColor();
	
	//::::13- solidFill (Solid Fill) 5.1.10.54
	private PPTXColor solidFill = new PPTXColor();
	
	
	private String TableCellWidth = "";
	private String TableCellHeight = "";
	
	/**
	 * @return Returns the anchor.
	 */
	public String getAnchor() {
		return anchor;
	}
	/**
	 * @param anchor The anchor to set.
	 */
	public void setAnchor(String anchor) {
		this.anchor = anchor;
	}
	/**
	 * @return Returns the anchorCtr.
	 */
	public boolean isAnchorCtr() {
		return anchorCtr;
	}
	/**
	 * @param anchorCtr The anchorCtr to set.
	 */
	public void setAnchorCtr(boolean anchorCtr) {
		this.anchorCtr = anchorCtr;
	}
	/**
	 * @return Returns the bgColor.
	 */
	public PPTXColor getBgColor() {
		return bgColor;
	}
	/**
	 * @param bgColor The bgColor to set.
	 */
	public void setBgColor(PPTXColor bgColor) {
		this.bgColor = bgColor;
	}
	/**
	 * @return Returns the fgColor.
	 */
	public PPTXColor getFgColor() {
		return fgColor;
	}
	/**
	 * @param fgColor The fgColor to set.
	 */
	public void setFgColor(PPTXColor fgColor) {
		this.fgColor = fgColor;
	}
	/**
	 * @return Returns the groupFill.
	 */
	public boolean isGroupFill() {
		return groupFill;
	}
	/**
	 * @param groupFill The groupFill to set.
	 */
	public void setGroupFill(boolean groupFill) {
		this.groupFill = groupFill;
	}
	/**
	 * @return Returns the horzOverflow.
	 */
	public String getHorzOverflow() {
		return horzOverflow;
	}
	/**
	 * @param horzOverflow The horzOverflow to set.
	 */
	public void setHorzOverflow(String horzOverflow) {
		this.horzOverflow = horzOverflow;
	}
	/**
	 * @return Returns the marB.
	 */
	public String getMarB() {
		return marB;
	}
	/**
	 * @param marB The marB to set.
	 */
	public void setMarB(String marB) {
		this.marB = marB;
	}
	/**
	 * @return Returns the marL.
	 */
	public String getMarL() {
		return marL;
	}
	/**
	 * @param marL The marL to set.
	 */
	public void setMarL(String marL) {
		this.marL = marL;
	}
	/**
	 * @return Returns the marR.
	 */
	public String getMarR() {
		return marR;
	}
	/**
	 * @param marR The marR to set.
	 */
	public void setMarR(String marR) {
		this.marR = marR;
	}
	/**
	 * @return Returns the marT.
	 */
	public String getMarT() {
		return marT;
	}
	/**
	 * @param marT The marT to set.
	 */
	public void setMarT(String marT) {
		this.marT = marT;
	}
	/**
	 * @return Returns the noFill.
	 */
	public boolean isNoFill() {
		return noFill;
	}
	/**
	 * @param noFill The noFill to set.
	 */
	public void setNoFill(boolean noFill) {
		this.noFill = noFill;
	}
	/**
	 * @return Returns the pptxBlipFill.
	 */
	public PPTXBlipFill getPptxBlipFill() {
		return pptxBlipFill;
	}
	/**
	 * @param pptxBlipFill The pptxBlipFill to set.
	 */
	public void setPptxBlipFill(PPTXBlipFill pptxBlipFill) {
		this.pptxBlipFill = pptxBlipFill;
	}
	/**
	 * @return Returns the pptxBottomBorderLineProperties.
	 */
	public PPTXOutline getPptxBottomBorderLineProperties() {
		return pptxBottomBorderLineProperties;
	}
	/**
	 * @param pptxBottomBorderLineProperties The pptxBottomBorderLineProperties to set.
	 */
	public void setPptxBottomBorderLineProperties(
			PPTXOutline pptxBottomBorderLineProperties) {
		this.pptxBottomBorderLineProperties = pptxBottomBorderLineProperties;
	}
	/**
	 * @return Returns the pptxBottomLToTopRightBorderProperties.
	 */
	public PPTXOutline getPptxBottomLToTopRightBorderProperties() {
		return pptxBottomLToTopRightBorderProperties;
	}
	/**
	 * @param pptxBottomLToTopRightBorderProperties The pptxBottomLToTopRightBorderProperties to set.
	 */
	public void setPptxBottomLToTopRightBorderProperties(
			PPTXOutline pptxBottomLToTopRightBorderProperties) {
		this.pptxBottomLToTopRightBorderProperties = pptxBottomLToTopRightBorderProperties;
	}
	/**
	 * @return Returns the pptxExtensionList.
	 */
	public PPTXExtensionList getPptxExtensionList() {
		return pptxExtensionList;
	}
	/**
	 * @param pptxExtensionList The pptxExtensionList to set.
	 */
	public void setPptxExtensionList(PPTXExtensionList pptxExtensionList) {
		this.pptxExtensionList = pptxExtensionList;
	}
	/**
	 * @return Returns the pptxGradientFill.
	 */
	public PPTXGradientFill getPptxGradientFill() {
		return pptxGradientFill;
	}
	/**
	 * @param pptxGradientFill The pptxGradientFill to set.
	 */
	public void setPptxGradientFill(PPTXGradientFill pptxGradientFill) {
		this.pptxGradientFill = pptxGradientFill;
	}
	/**
	 * @return Returns the pptxLeftBorderLineProperties.
	 */
	public PPTXOutline getPptxLeftBorderLineProperties() {
		return pptxLeftBorderLineProperties;
	}
	/**
	 * @param pptxLeftBorderLineProperties The pptxLeftBorderLineProperties to set.
	 */
	public void setPptxLeftBorderLineProperties(
			PPTXOutline pptxLeftBorderLineProperties) {
		this.pptxLeftBorderLineProperties = pptxLeftBorderLineProperties;
	}
	/**
	 * @return Returns the pptxTopBorderLineProperties.
	 */
	public PPTXOutline getPptxTopBorderLineProperties() {
		return pptxTopBorderLineProperties;
	}
	/**
	 * @param pptxTopBorderLineProperties The pptxTopBorderLineProperties to set.
	 */
	public void setPptxTopBorderLineProperties(
			PPTXOutline pptxTopBorderLineProperties) {
		this.pptxTopBorderLineProperties = pptxTopBorderLineProperties;
	}
	/**
	 * @return Returns the pptxTopLToBottomRightBorderProperties.
	 */
	public PPTXOutline getPptxTopLToBottomRightBorderProperties() {
		return pptxTopLToBottomRightBorderProperties;
	}
	/**
	 * @param pptxTopLToBottomRightBorderProperties The pptxTopLToBottomRightBorderProperties to set.
	 */
	public void setPptxTopLToBottomRightBorderProperties(
			PPTXOutline pptxTopLToBottomRightBorderProperties) {
		this.pptxTopLToBottomRightBorderProperties = pptxTopLToBottomRightBorderProperties;
	}
	/**
	 * @return Returns the presetPattern.
	 */
	public String getPresetPattern() {
		return presetPattern;
	}
	/**
	 * @param presetPattern The presetPattern to set.
	 */
	public void setPresetPattern(String presetPattern) {
		this.presetPattern = presetPattern;
	}
	/**
	 * @return Returns the solidFill.
	 */
	public PPTXColor getSolidFill() {
		return solidFill;
	}
	/**
	 * @param solidFill The solidFill to set.
	 */
	public void setSolidFill(PPTXColor solidFill) {
		this.solidFill = solidFill;
	}
	/**
	 * @return Returns the vert.
	 */
	public String getVert() {
		return vert;
	}
	/**
	 * @param vert The vert to set.
	 */
	public void setVert(String vert) {
		this.vert = vert;
	}
	/**
	 * @return Returns the tableCellHeight.
	 */
	public String getTableCellHeight() {
		return TableCellHeight;
	}
	/**
	 * @param tableCellHeight The tableCellHeight to set.
	 */
	public void setTableCellHeight(String tableCellHeight) {
		TableCellHeight = tableCellHeight;
	}
	/**
	 * @return Returns the tableCellWidth.
	 */
	public String getTableCellWidth() {
		return this.TableCellWidth;
	}
	/**
	 * @param tableCellWidth The tableCellWidth to set.
	 */
	public void setTableCellWidth(String tableCellWidth) {
		this.TableCellWidth = tableCellWidth;
	}
	/**
     * Description: Set attributes of Table Cell Properties instance by using incoming tcPr node
     * <p>
     * @param attributes instance of NamedNodeMap - its input will be array of attributes of tcPr node
     * @author Shabana Majeed
     */

	public void populateAttributeValues(NamedNodeMap attributes)
	{

		// tcPr (Table Cell Properties) 5.1.6.15
		
//		private String anchor = "";			//Defines the alignment of the text vertically within the cell.
//		private boolean anchorCtr = false;  //When this attribute is on, 1 or true, it modifies the anchor attribute
//		private String horzOverflow = "";   //Specifies the clipping behavior of the cell.
//		private String marB = "";    //Specifies the bottom margin of the cell      
//		private String marL = "";	 //Specifies the left margin of the cell
//		private String marR = "";    //Specifies the right margin of the cell
//		private String marT = "";    //Specifies the top margin of the cell
//		private String vert = "";	 //Defines the text direction within the cell
		
		if(attributes != null && attributes.getLength() > 0)
		{
			for(int at = 0; at < attributes.getLength(); at++)
			{
				Node attribute = attributes.item(at);
				String attributeName = attribute.getNodeName();
				String attributeValue = attribute.getNodeValue();
				
				// anchor (Anchor)
				if(attributeName.equalsIgnoreCase("anchor"))
				{
					this.setAnchor(attributeValue);
				}
				
				// anchorCtr (Anchor Center)
				else if(attributeName.equalsIgnoreCase("anchorCtr"))
				{
					if(attributeValue.equals("1"))
					{
						this.setAnchorCtr(true);
					}
				}
				
				// horzOverflow (Horizontal Overflow)
				else if(attributeName.equalsIgnoreCase("horzOverflow"))
				{
					this.setHorzOverflow(attributeValue);
				}
				
				// marB (Bottom Margin)
				else if(attributeName.equalsIgnoreCase("marB"))
				{
					this.setMarB(attributeValue);
				}
				
				// marL (Left Margin)
				else if(attributeName.equalsIgnoreCase("marL"))
				{
					this.setMarL(attributeValue);
				}
				
				// marR (Right Margin)
				else if(attributeName.equalsIgnoreCase("marR"))
				{
					this.setMarR(attributeValue);
				}
				
				// marT (Top Margin)
				else if(attributeName.equalsIgnoreCase("marT"))
				{
					this.setMarT(attributeValue);
				}
				
				// vert (Text Direction)
				else if(attributeName.equalsIgnoreCase("vert"))
				{
					this.setVert(attributeValue);
				}
			}
		}
	}
	
	/**
     * Description: Parse children of Table Cell Properties instance by using incoming tcPr node
     * <p>
     * @param children instance of NodeList - its input will be array of children of tcPr node
     * @author Shabana Majeed
     */	
	
	public void parseChildren(NodeList children)
	{
		
		// tcPr (Table Cell Properties) 5.1.6.15
		
		if(children.getLength() > 0)
		{
			for(int chIndex = 0; chIndex < children.getLength(); chIndex++)
			{
				Node child = children.item(chIndex);
				String childName = child.getNodeName();
				
				// blipFill (Picture Fill) 5.1.10.14
				// private PPTXBlipFill pptxBlipFill = new PPTXBlipFill();
				if(childName.contains("blipFill"))
				{
					this.getPptxBlipFill().populateAttributeValues(child.getAttributes());
					this.getPptxBlipFill().parseChildren(child.getChildNodes());
				}
				
				// cell3D (Cell 3-D) 5.1.6.1
				else if(childName.contains("cell3D"))
				{

				}
				
				// extLst (Extension List) 5.1.2.1.15
				// private PPTXExtensionList pptxExtensionList = new PPTXExtensionList();
				else if(childName.contains("extLst"))
				{
					this.getPptxExtensionList().populateAttributeValues(child.getAttributes());
					this.getPptxExtensionList().parseChildren(child.getChildNodes());
				}
				
				// gradFill (Gradient Fill) 5.1.10.33
				// private PPTXGradientFill pptxGradientFill = new PPTXGradientFill();
				else if(childName.contains("gradFill"))
				{
				//	this.getPptxGradientFill().populateAttributeValues(child.getAttributes());
				//	this.getPptxGradientFill().parseChildren(child.getChildNodes());
				}
				
				// grpFill (Group Fill) 5.1.10.35
				// private boolean groupFill = false;//This element specifies a group fill.
				else if(childName.contains("grpFill"))
				{
					this.setGroupFill(true);
				}				
				
				// lnB (Bottom Border Line Properties) 5.1.6.3
				// private PPTXOutline pptxBottomBorderLineProperties = new PPTXOutline();
				else if(childName.contains("lnB"))
				{
					this.getPptxBottomBorderLineProperties().populateAttributeValues(child.getAttributes());
					this.getPptxBottomBorderLineProperties().parseChildren(child.getChildNodes());
				}
				
				// lnBlToTr (Bottom-Left to Top-Right Border Line Properties) 5.1.6.4
				// private PPTXOutline pptxBottomLToTopRightBorderProperties = new PPTXOutline();
				else if(childName.contains("lnBlToTr"))
				{
					this.getPptxBottomLToTopRightBorderProperties().populateAttributeValues(child.getAttributes());
					this.getPptxBottomLToTopRightBorderProperties().parseChildren(child.getChildNodes());
				}
				
				// lnL (Left Border Line Properties) 5.1.6.5
				// private PPTXOutline pptxLeftBorderLineProperties = new PPTXOutline();
				else if(childName.contains("lnL"))
				{
					this.getPptxLeftBorderLineProperties().populateAttributeValues(child.getAttributes());
					this.getPptxLeftBorderLineProperties().parseChildren(child.getChildNodes());
				}
				
				// lnTlToBr (Top-Left to Bottom-Right Border Line Properties) 5.1.6.8
				// private PPTXOutline pptxTopLToBottomRightBorderProperties = new PPTXOutline();
				else if(childName.contains("lnTlToBr"))
				{
					this.getPptxTopLToBottomRightBorderProperties().populateAttributeValues(child.getAttributes());
					this.getPptxTopLToBottomRightBorderProperties().parseChildren(child.getChildNodes());
				}
				
				// noFill (No Fill) 5.1.10.44
				// private boolean noFill = false;//This element specifies that no fill will be applied to the parent element.
				else if(childName.contains("noFill"))
				{
					this.setNoFill(true);
				}
	
				// pattFill (pattern Fill) 5.1.10.47
				// private String presetPattern = ""; //Specifies one of a set of preset patterns to fill the object.
				// private PPTXColor bgColor = new PPTXColor();
				// private PPTXColor fgColor = new PPTXColor();
				else if(childName.contains("pattFill"))
				{
					NamedNodeMap pattFillAttributes = child.getAttributes();
					
					if (pattFillAttributes != null)
					{
						// prst (Preset Pattern)
                        if(pattFillAttributes.getNamedItem("prst") != null)
                        {
							this.setPresetPattern(pattFillAttributes.getNamedItem("prst").getNodeValue());
                        }
					}
					
					NodeList pattFillChildren = child.getChildNodes();
					
					if(pattFillChildren.getLength() > 0)
					{
						for(int childIndex = 0; childIndex < pattFillChildren.getLength(); childIndex++)
						{
							Node pattFillchild = pattFillChildren.item(childIndex);
							String pattFillchildName = pattFillchild.getNodeName();
							
							// bgClr (Background Color) 5.1.10.10
							if(pattFillchildName.contains("bgClr"))
							{
								this.getBgColor().parseChildren(pattFillchild.getChildNodes());
							}
							
							// fgClr (Foreground Color) 5.1.10.27
							else if(pattFillchildName.contains("fgClr"))
							{
								this.getFgColor().parseChildren(pattFillchild.getChildNodes());
							}
						}
					}
				}
				
				// solidFill (Solid Fill) 5.1.10.54
				// private PPTXColor solidFill = new PPTXColor();
				else if(childName.contains("solidFill"))
				{
					this.getSolidFill().parseChildren(child.getChildNodes());
				}				
			}
		}
	}	
	

	/**
     * Description: Calculate the exact position of cell by using row and column details
     * <p>
     * @param int (row number,column width and row Height) 
     * @author Shabana Majeed
     */	
	
	public void setCellMargins(int rowNum, int colWidth, int rowHeight)
	{		
   	
		this.setMarL(Integer.toString(context.getPrevColumnWidth() + context.getPrevCellX()));
	     	this.setTableCellWidth(Integer.toString(colWidth));
	     	this.setTableCellHeight(Integer.toString(rowHeight));
	     	
	     	if(context.isInSameRow())
	     	{
	     		this.setMarT(Integer.toString(context.getPrevRowY() + context.getPrevRowHeight()));	     	
	     		context.setPrevRowY(Integer.parseInt(this.getMarT()));
	     		context.setPrevCellY(Integer.parseInt(this.getMarT()));
	     	}	     	
	     	else
	     	{
	     		this.setMarT(Integer.toString(context.getPrevCellY()));
	     	}
	     	
	     	context.setPrevCellX(Integer.parseInt(this.getMarL()));
	     	context.setPrevCellY(Integer.parseInt(this.getMarT()));
	     	context.setPrevColumnWidth(colWidth);
	}

}
