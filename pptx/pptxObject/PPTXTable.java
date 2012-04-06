package pptx.pptxObject;

import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;



/**
 * @author Sheraz Ahmed
 *
 */
public class PPTXTable 
{
	private PPTXContext context = PPTXContext.getInstance();
	
	
	// tbl (Table) 5.1.6.11
	
	// Children
	//::::1- tblGrid (Table Grid) 5.1.6.12
	// Children
	// gridCol (Table Grid Column) 5.1.6.2
	
	// Children
	// extLst (Extension List) 5.1.2.1.15
	private PPTXExtensionList pptxExtensionList = new PPTXExtensionList();


	//::::2- tblPr (Table Properties) 5.1.6.13
	private PPTXTableProperties tableProperties = new PPTXTableProperties();
	
	//::::3- tr (Table Row) 5.1.6.16
    private String height = "";
	// Children
	// i-extLst (Extension List) 5.1.2.1.15
    private PPTXExtensionList pptxTrExtensionList = new PPTXExtensionList();
 
		
    // ii-tc (Table Cell) 5.1.6.14
    
	private List<PPTXTableCell> pptxTableCellList = new ArrayList<PPTXTableCell>();

	public PPTXExtensionList getPptxExtensionList() {
		return pptxExtensionList;
	}

	/**
	 * @return Returns the pptxTableCellList.
	 */
	public List<PPTXTableCell> getPptxTableCellList() {
		return pptxTableCellList;
	}

	/**
	 * @param pptxTableCellList The pptxTableCellList to set.
	 */
	public void setPptxTableCellList(List<PPTXTableCell> pptxTableCellList) {
		this.pptxTableCellList = pptxTableCellList;
	}

	/**
	 * @param pptxExtensionList The pptxExtensionList to set.
	 */
	public void setPptxExtensionList(PPTXExtensionList pptxExtensionList) {
		this.pptxExtensionList = pptxExtensionList;
	}
	/**
	 * @return Returns the pptxTrExtensionList.
	 */
	public PPTXExtensionList getPptxTrExtensionList() {
		return pptxTrExtensionList;
	}

	/**
	 * @param pptxTrExtensionList The pptxTrExtensionList to set.
	 */
	public void setPptxTrExtensionList(PPTXExtensionList pptxTrExtensionList) {
		this.pptxTrExtensionList = pptxTrExtensionList;
	}

	/**
	 * @return Returns the tableProperties.
	 */
	public PPTXTableProperties getTableProperties() {
		return tableProperties;
	}

	/**
	 * @param tableProperties The tableProperties to set.
	 */
	public void setTableProperties(PPTXTableProperties tableProperties) {
		this.tableProperties = tableProperties;
	}

	/**
	 * @return Returns the height.
	 */
	public String getHeight() {
		return height;
	}

	/**
	 * @param height The height to set.
	 */
	public void setHeight(String height) {
		this.height = height;
	}
	/**
     * Description: Parse children of Table instance by using incoming tbl node
     * <p>
     * @param children instance of NodeList - its input will be array of children of tbl node
     * @author Shabana Majeed
     */	
	
    public void parseChildren(NodeList children)
	{
		// tbl (Table) 5.1.6.11
		
		if(children.getLength() > 0)
		{
			for(int chIndex = 0; chIndex < children.getLength(); chIndex++)
			{
				Node child = children.item(chIndex);
				String childName = child.getNodeName();
				
				// tblGrid (Table Grid) 5.1.6.12
				if(childName.contains("tblGrid"))
				{
					if (child.hasChildNodes())
					{
						NodeList tblGridchildren = child.getChildNodes();
						for(int tblGridIndex = 0; tblGridIndex < tblGridchildren.getLength(); tblGridIndex++)
						{
							Node tblGridChild = tblGridchildren.item(tblGridIndex);
							String tblGridChildName = tblGridChild.getNodeName();
							
							// gridCol (Table Grid Column) 5.1.6.2
							if(tblGridChildName.contains("gridCol"))
							{
								context.getColumnWidths().add(tblGridIndex,
										tblGridChild.getAttributes().getNamedItem("w").getNodeValue());
								
								if (tblGridChild.hasChildNodes())
								{
									// extLst (Extension List) 5.1.2.1.15
									Node extLstChild = tblGridChild.getFirstChild();
									this.getPptxExtensionList().populateAttributeValues(extLstChild.getAttributes());
									this.getPptxExtensionList().parseChildren(extLstChild.getChildNodes());
								}
							}
						}
					}
				}
				
				// tblPr (Table Properties) 5.1.6.13
				else if(childName.contains("tblPr"))
				{
					this.getTableProperties().populateAttributeValues(child.getAttributes());
					this.getTableProperties().parseChildren(child.getChildNodes());
				}
				
				// tr (Table Row) 5.1.6.16
				else if(childName.contains("tr"))
				{
					context.setInSameRow(true);
					context.setRowCount(context.getRowCount()+1);
					context.setCellCount(0);
					context.setPrevCellX(0);
					context.setPrevColumnWidth(0);
					
					String cellHeight = child.getAttributes().getNamedItem("h").getNodeValue();
					
					NodeList trChildren = child.getChildNodes();
					if(trChildren.getLength() > 0)
					{
						for(int childIndex = 0; childIndex < trChildren.getLength(); childIndex++)
						{
							Node trChild = trChildren.item(childIndex);
							String trChildName = trChild.getNodeName();
							
							// extLst (Extension List) 5.1.2.1.15
					    	// private PPTXExtensionList pptxTcExtensionList = new PPTXExtensionList();
							if(trChildName.contains("extLst"))
							{
								this.getPptxExtensionList().populateAttributeValues(trChild.getAttributes());
								this.getPptxExtensionList().parseChildren(trChild.getChildNodes());
							}
							
							// tc (Table Cell) 5.1.6.14
							else if(trChildName.contains("tc"))
							{							
								context.setCellCount(context.getCellCount() + 1);
								PPTXTableCell pptxTableCell = new PPTXTableCell();
								pptxTableCell.populateAttributeValues(trChild.getAttributes());
								context.getRowHeights().add (context.getRowCount() - 1, cellHeight);
								pptxTableCell.parseChildren(trChild.getChildNodes());
								context.setInSameRow(false);
								this.getPptxTableCellList().add(pptxTableCell);								
								
							}//end of tc block
						}	
					}
					context.setPrevRowHeight(Integer.parseInt(context.getRowHeights().get(context.getRowCount() - 1)));					
				} // end of tr block
			}
		}
	}//end of method block
}