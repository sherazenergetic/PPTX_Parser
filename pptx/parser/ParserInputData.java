/**
 * 
 */
package pptx.parser;

import java.io.InputStream;
import java.util.HashMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.w3c.dom.Element;

import pptx.pptxObject.PPTXGraphicFrame;
import pptx.pptxObject.PPTXShape;

/**
 * @author sheraz.ahmed
 *
 */
public class ParserInputData {
	
	private ZipEntry relsZipEntry = null; //_rels/.rels entry

	private ZipEntry coreZipEntry = null;//docProps/core.xml
	
	private ZipEntry appZipEntry = null; //docProps/app.xml 	
	
	private ZipEntry presentationEntry = null;

	private HashMap<String, String> presentationRelationships = null;
	
	private ZipEntry slide = null;
	
	private ZipEntry slideLayout = null;
	
	private HashMap<String, String> slideRelationships = null;
	
	private PPTXShape shapeElement = null;
	
	private PPTXGraphicFrame graphicFrameElement = null;
	
	private PPTXShape shapeElementOfLayout = null;

	private Element txBodyElement = null;
	
	private Element txBodyElementOfLayout = null;

	private Element PPTXParaElement = null;
	
	private Element pptxParaElementOfLayout = null;
	
	private ZipEntry slideMasterZipEntry = null;
	
	private ZipEntry masterSlideThemeZipEntry = null;
	
	private HashMap<String, InputStream> ImagesStreamsMap = null;

	/**
	 * 
	 * @return coreZipEntry for docProps/core.xml
	 */
	public ZipEntry getCoreZipEntry() 
	{
		return coreZipEntry;
	}

	/**
	 * 
	 * @param coreZipEntry for docProps/core.xml
	 */
	public void setCoreZipEntry(ZipEntry coreZipEntry) 
	{
		this.coreZipEntry = coreZipEntry;
	}

	/**
	 * 
	 * @return rels ZipEntry for _rels/.rels
	 */
	public ZipEntry getRelsZipEntry() 
	{
		return relsZipEntry;
	}

	/**
	 * 
	 * @param rels ZipEntry for _rels/.rels
	 */
	public void setRelsZipEntry(ZipEntry relsZipEntry) 
	{
		this.relsZipEntry = relsZipEntry;
	}
	
	/**
	 * 
	 * @return appZipEntry for docProps/app.xml
	 */
	public ZipEntry getAppZipEntry() 
	{
		return appZipEntry;
	}

	/**
	 * 
	 * @param appZipEntry for docProps/app.xml
	 */
	public void setAppZipEntry(ZipEntry appZipEntry) 
	{
		this.appZipEntry = appZipEntry;
	}

	/**
	 * 
	 * @return presentationEntry
	 */
	public ZipEntry getPresentationEntry() 
	{
		return presentationEntry;
	}

	/**
	 * 
	 * @param presentationEntry
	 */
	public void setPresentationEntry(ZipEntry presentationEntry) 
	{
		this.presentationEntry = presentationEntry;
	}

	/**
	 * 
	 * @return presentationRelationships
	 */
	public HashMap<String, String> getPresentationRelationships() 
	{
		return presentationRelationships;
	}

	/**
	 * 
	 * @param presentationRelationships
	 */
	public void setPresentationRelationships(
			HashMap<String, String> presentationRelationships) 
	{
		this.presentationRelationships = presentationRelationships;
	}

	/**
	 * 
	 * @return slide instance of ZipEntry
	 */
	public ZipEntry getSlide() 
	{
		return slide;
	}

	/**
	 * 
	 * @param slide instance of ZipEntry
	 */
	public void setSlide(ZipEntry slide) 
	{
		this.slide = slide;
	}

	/**
	 * 
	 * @return slide Layout ZipEntry
	 */
	public ZipEntry getSlideLayout() 
	{
		return slideLayout;
	}

	/**
	 * 
	 * @param slide Layout ZipEntry
	 */
	public void setSlideLayout(ZipEntry slideLayout) 
	{
		this.slideLayout = slideLayout;
	}

	/**
	 * 
	 * @return txBodyElement
	 */
	public Element getTXBodyElement() 
	{
		return txBodyElement;
	}

	/**
	 * 
	 * @param txBodyElement
	 */
	public void setTXBodyElement(Element txBodyElement) 
	{
		this.txBodyElement = txBodyElement;
	}

	/**
	 * 
	 * @return txBodyElementOfLayout
	 */
	public Element getTXBodyElementOfLayout() 
	{
		return txBodyElementOfLayout;
	}

	/**
	 * 
	 * @param txBodyElementOfLayout
	 */
	public void setTXBodyElementOfLayout(Element txBodyElementOfLayout) 
	{
		this.txBodyElementOfLayout = txBodyElementOfLayout;
	}

	/**
	 * 
	 * @return PPTXParaElement
	 */
	public Element getPPTXParaElement() 
	{
		return PPTXParaElement;
	}

	/**
	 * 
	 * @param pPTXParaElement
	 */
	public void setPPTXParaElement(Element PPTXParaElement) 
	{
		this.PPTXParaElement = PPTXParaElement;
	}

	/**
	 * 
	 * @return pptxParaElementOfLayout
	 */
	public Element getPPTXParaElementOfLayout() {
		return pptxParaElementOfLayout;
	}

	/**
	 * 
	 * @param pptxParaElementOfLayout
	 */
	public void setPPTXParaElementOfLayout(Element pptxParaElementOfLayout) {
		this.pptxParaElementOfLayout = pptxParaElementOfLayout;
	}

	/**
	 * @return the slideMasterZipEntry
	 */
	public ZipEntry getSlideMasterZipEntry() {
		return slideMasterZipEntry;
	}

	/**
	 * @param slideMasterZipEntry the slideMasterZipEntry to set
	 */
	public void setSlideMasterZipEntry(ZipEntry slideMasterZipEntry) {
		this.slideMasterZipEntry = slideMasterZipEntry;
	}

	/**
	 * @return the masterSlideThemeZipEntry
	 */
	public ZipEntry getMasterSlideThemeZipEntry() {
		return masterSlideThemeZipEntry;
	}

	/**
	 * @param masterSlideThemeZipEntry the masterSlideThemeZipEntry to set
	 */
	public void setMasterSlideThemeZipEntry(ZipEntry masterSlideThemeZipEntry) {
		this.masterSlideThemeZipEntry = masterSlideThemeZipEntry;
	}

	/**
	 * @return the shapeElement
	 */
	public PPTXShape getShapeElement() {
		return shapeElement;
	}

	/**
	 * @param shapeElement the shapeElement to set
	 */
	public void setShapeElement(PPTXShape shapeElement) {
		this.shapeElement = shapeElement;
	}

	/**
	 * @return the shapeElementOfLayout
	 */
	public PPTXShape getShapeElementOfLayout() {
		return shapeElementOfLayout;
	}

	/**
	 * @param shapeElementOfLayout the shapeElementOfLayout to set
	 */
	public void setShapeElementOfLayout(PPTXShape shapeElementOfLayout) {
		this.shapeElementOfLayout = shapeElementOfLayout;
	}

	/**
	 * @return the imagesStreamsMap
	 */
	public HashMap<String, InputStream> getImagesStreamsMap() {
		return ImagesStreamsMap;
	}

	/**
	 * @param imagesStreamsMap the imagesStreamsMap to set
	 */
	public void setImagesStreamsMap(HashMap<String, InputStream> imagesStreamsMap) {
		ImagesStreamsMap = imagesStreamsMap;
	}
	
	/**
	 * @return Returns the graphicFrameElement.
	 */
	public PPTXGraphicFrame getGraphicFrameElement() {
		return graphicFrameElement;
	}

	/**
	 * @param graphicFrameElement The graphicFrameEleement to set.
	 */
	public void setGraphicFrameElement(PPTXGraphicFrame graphicFrameEleement) {
		this.graphicFrameElement = graphicFrameEleement;
	}

	/**
	 * @return Returns the slideRelationships.
	 */
	public HashMap<String, String> getSlideRelationships() {
		return slideRelationships;
	}

	/**
	 * @param slideRelationships The slideRelationships to set.
	 */
	public void setSlideRelationships(HashMap<String, String> slideRelationships) {
		this.slideRelationships = slideRelationships;
	}

	
}
