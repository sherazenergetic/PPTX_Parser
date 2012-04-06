
package pptx.pptxObject;

import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import pptx.context.PPTXContext;

/**
 * @author Sheraz Ahmed
 *
 */
public class PPTXTableCell {
	
	private PPTXContext context = PPTXContext.getInstance();
	
	
	    // tc (Table Cell) 5.1.6.14
	    private int gridSpan = 0;        // Specifies the number of columns that a merged cell spans
	    private boolean hMerge = false;  // When this attribute is set to on, 1 or true, then this table cell is to be merged with the previous horizontal table cell when the table is created
	    private int rowSpan = 0;    	 // Specifies the number of rows that a merged cell spans    
	    private boolean vMerge = false;  // When this attribute is set to on, 1 or true, then this table cell is to be merged with the previous vertical table cell when the table is created
	    // Children
	    // a-extLst (Extension List) 5.1.2.1.15
		private PPTXExtensionList pptxTcExtensionList = new PPTXExtensionList();
		// b-tcPr (Table Cell Properties) 5.1.6.15
		private PPTXTableCellProperties pptxTableCellProperties = new PPTXTableCellProperties();
		// c-txBody (Shape Text Body) 5.1.2.1.40
		//children
		// bodyPr (Body Properties) 5.1.5.1.1
		private PPTXBodyProperties pptxBodyProperties = new PPTXBodyProperties();
		//lstStyle (Text List Styles) 5.1.5.4.12
		private PPTXListTextStyles pptxListTextStyles = new PPTXListTextStyles(); 
		//p (Text Paragraphs) 5.1.5.2.6
		private List<PPTXTextParagraph> pptxTextParagraphList = new ArrayList<PPTXTextParagraph>();
		private String tableCellHeight = "";

		/**
		 * @return Returns the gridSpan.
		 */
		public int getGridSpan() {
			return gridSpan;
		}


		/**
		 * @param gridSpan The gridSpan to set.
		 */
		public void setGridSpan(int gridSpan) {
			this.gridSpan = gridSpan;
		}


		/**
		 * @return Returns the hMerge.
		 */
		public boolean isHMerge() {
			return hMerge;
		}


		/**
		 * @param merge The hMerge to set.
		 */
		public void setHMerge(boolean merge) {
			hMerge = merge;
		}


		/**
		 * @return Returns the pptxBodyProperties.
		 */
		public PPTXBodyProperties getPptxBodyProperties() {
			return pptxBodyProperties;
		}


		/**
		 * @param pptxBodyProperties The pptxBodyProperties to set.
		 */
		public void setPptxBodyProperties(PPTXBodyProperties pptxBodyProperties) {
			this.pptxBodyProperties = pptxBodyProperties;
		}


		/**
		 * @return Returns the pptxListTextStyles.
		 */
		public PPTXListTextStyles getPptxListTextStyles() {
			return pptxListTextStyles;
		}


		/**
		 * @param pptxListTextStyles The pptxListTextStyles to set.
		 */
		public void setPptxListTextStyles(PPTXListTextStyles pptxListTextStyles) {
			this.pptxListTextStyles = pptxListTextStyles;
		}


		/**
		 * @return Returns the pptxTableCellProperties.
		 */
		public PPTXTableCellProperties getPptxTableCellProperties() {
			return pptxTableCellProperties;
		}


		/**
		 * @param pptxTableCellProperties The pptxTableCellProperties to set.
		 */
		public void setPptxTableCellProperties(
				PPTXTableCellProperties pptxTableCellProperties) {
			this.pptxTableCellProperties = pptxTableCellProperties;
		}


		/**
		 * @return Returns the pptxTcExtensionList.
		 */
		public PPTXExtensionList getPptxTcExtensionList() {
			return pptxTcExtensionList;
		}


		/**
		 * @param pptxTcExtensionList The pptxTcExtensionList to set.
		 */
		public void setPptxTcExtensionList(PPTXExtensionList pptxTcExtensionList) {
			this.pptxTcExtensionList = pptxTcExtensionList;
		}


		/**
		 * @return Returns the pptxTextParagraphList.
		 */
		public List<PPTXTextParagraph> getPptxTextParagraphList() {
			return pptxTextParagraphList;
		}


		/**
		 * @param pptxTextParagraphList The pptxTextParagraphList to set.
		 */
		public void setPptxTextParagraphList(
				List<PPTXTextParagraph> pptxTextParagraphList) {
			this.pptxTextParagraphList = pptxTextParagraphList;
		}


		/**
		 * @return Returns the rowSpan.
		 */
		public int getRowSpan() {
			return rowSpan;
		}


		/**
		 * @param rowSpan The rowSpan to set.
		 */
		public void setRowSpan(int rowSpan) {
			this.rowSpan = rowSpan;
		}


		/**
		 * @return Returns the vMerge.
		 */
		public boolean isVMerge() {
			return vMerge;
		}


		/**
		 * @param merge The vMerge to set.
		 */
		public void setVMerge(boolean merge) {
			vMerge = merge;
		}
		



		/**
	     * Description: Set attributes of Table Cell instance by using incoming tc node
	     * <p>
	     * @param attributes instance of NamedNodeMap - its input will be array of attributes of tc node
	     * @author Shabana Majeed
	     */

		public void populateAttributeValues(NamedNodeMap attributes)
		{
//			 tc (Table Cell) 5.1.6.14
//		    private int gridSpan = 0;        // Specifies the number of columns that a merged cell spans
//		    private boolean hMerge = false;  // When this attribute is set to on, 1 or true, then this table cell is to be merged with the previous horizontal table cell when the table is created
//		    private int rowSpan = 0;    	 // Specifies the number of rows that a merged cell spans    
//		    private boolean vMerge = false;  // When this attribute is set to on, 1 or true, then this table cell is to be merged with the previous vertical table cell when the table is created
			
					
			if(attributes != null && attributes.getLength() > 0)
			{
				for(int at = 0; at < attributes.getLength(); at++)
				{
					Node attribute = attributes.item(at);
					String attributeName = attribute.getNodeName();
					String attributeValue = attribute.getNodeValue();
					
					// gridSpan (Grid Span) 
					if(attributeName.equalsIgnoreCase("gridSpan"))
					{
						if(attributeValue != null && attributeValue.trim().length() > 0)
						{
							this.setGridSpan(Integer.parseInt(attributeValue));
	                    }
					}
					
					// hMerge (Horizontal Merge)
					else if(attributeName.equalsIgnoreCase("hMerge"))
					{
						if(attributeValue.equals("1"))
						{
							this.setHMerge(true);
						}
					}
					
					// rowSpan (Row Span)
					else if(attributeName.equalsIgnoreCase("rowSpan"))
					{
						if(attributeValue != null && attributeValue.trim().length() > 0)
						{
							this.setRowSpan(Integer.parseInt(attributeValue));
	                    }
					}
					
					// vMerge (Vertical Merge)
					else if(attributeName.equalsIgnoreCase("vMerge"))
					{
						if(attributeValue.equals("1"))
						{
							this.setVMerge(true);
						}
					}
				}
			}
		}		
		/**
	     * Description: Parse children of Table Cell instance by using incoming tc node
	     * <p>
	     * @param children instance of NodeList - its input will be array of children of tc node
	     * @author Shabana Majeed
	     */
		
	    public void parseChildren(NodeList children)
		{
	    	//	tc (Table Cell) 5.1.6.14
			
			if(children.getLength() > 0)
			{
				for(int tcChildIndex = 0; tcChildIndex < children.getLength(); tcChildIndex++)
				{
					Node tcChild = children.item(tcChildIndex);
					String tcChildName = tcChild.getNodeName();
					
					// extLst (Extension List) 5.1.2.1.15
					// private PPTXExtensionList pptxTcExtensionList = new PPTXExtensionList();
					if(tcChildName.contains("extLst"))
					{
						this.getPptxTcExtensionList().populateAttributeValues(tcChild.getAttributes());
						this.getPptxTcExtensionList().parseChildren(tcChild.getChildNodes());
					}
					
					// tcPr (Table Cell Properties) 5.1.6.15
					else if(tcChildName.contains("tcPr"))
					{
						this.getPptxTableCellProperties().populateAttributeValues(tcChild.getAttributes());
						this.getPptxTableCellProperties().parseChildren(tcChild.getChildNodes());
						
						this.getPptxTableCellProperties().setCellMargins(context.getRowCount(),
								Integer.parseInt(context.getColumnWidths().get(context.getCellCount() - 1)),
								Integer.parseInt(context.getRowHeights().get(context.getRowCount() - 1)));
					}
					
					// txBody (shape text body) 4.4.1.47
					// private PPTXBodyProperties pptxBodyProperties = new PPTXBodyProperties();
					// private PPTXListTextStyles pptxListTextStyles = new PPTXListTextStyles(); 
					// private PPTXTextParagraph pptxTextParagraph = new PPTXTextParagraph();
					else if (tcChildName.contains("txBody"))
						
					{
						NodeList txtBodyChildren = tcChild.getChildNodes();
		
						for(int txtBodyChildIndex = 0; txtBodyChildIndex < txtBodyChildren.getLength(); txtBodyChildIndex++)
						{
							Node txtBodyChild = txtBodyChildren.item(txtBodyChildIndex);
							String txtBodyChildName = txtBodyChild.getNodeName();
		
							// bodyPr (Body Properties) 5.1.5.1.1
							if(txtBodyChildName.contains("bodyPr"))
							{
								this.getPptxBodyProperties().populateAttributeValues(txtBodyChild.getAttributes());
								this.getPptxBodyProperties().parseChildren(txtBodyChild.getChildNodes());
							}
							
							// lstStyle (Text List Styles) 5.1.5.4.12
							else if(txtBodyChildName.contains("lstStyle"))
							{
								this.getPptxListTextStyles().parseChildren(txtBodyChild.getChildNodes());
							}
		
							// p (Text Paragraphs) 5.1.5.2.6 
							else if(txtBodyChildName.contains("p"))
							{
								PPTXTextParagraph pptxTextParagraph = new PPTXTextParagraph();
								pptxTextParagraph.parseChildren(txtBodyChild.getChildNodes());
								this.getPptxTextParagraphList().add(pptxTextParagraph);
							}						
						}
					}
				} //end of for loop
			}
		}//end of method block	
}
