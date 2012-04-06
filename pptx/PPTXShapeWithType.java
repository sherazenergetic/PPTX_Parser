/**
 * 
 */
package pptx;

import pptx.pptxObject.PPTXGraphicFrame;
import pptx.pptxObject.PPTXShape;
import pptx.pptxObject.PPTXTable;

/**
 * @author sheraz.ahmed
 *
 */
public class PPTXShapeWithType 
{

	private String type = "";
	private PPTXShape pptxShape = null;
	private PPTXGraphicFrame pptxGraphicFrame = null;
	/**
	 * @return the type
	 */
	public String getType() {
		return type;
	}
	/**
	 * @param type the type to set
	 */
	public void setType(String type) {
		this.type = type;
	}
	/**
	 * @return the pptxShape
	 */
	public PPTXShape getPptxShape() {
		return pptxShape;
	}
	/**
	 * @param pptxShape the pptxShape to set
	 */
	public void setPptxShape(PPTXShape pptxShape) {
		this.pptxShape = pptxShape;
	}
	/**
	 * @return Returns the pptxGraphicFrame.
	 */
	public PPTXGraphicFrame getPptxGraphicFrame() {
		return pptxGraphicFrame;
	}
	/**
	 * @param pptxGraphicFrame The pptxGraphicFrame to set.
	 */
	public void setPptxGraphicFrame(PPTXGraphicFrame pptxGraphicFrame) {
		this.pptxGraphicFrame = pptxGraphicFrame;
	}
	
	
	
}
