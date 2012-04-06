
package pptx.context;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import ppXML.PPXMLDocument;
import ppXML.PPXMLObjectFactory;
import pptx.PPTXConstants;
import pptx.pptxObject.PPTXDefaultTextRunProperties;
import pptx.pptxObject.PPTXComplexScriptFont;
import pptx.pptxObject.PPTXEastAsianFont;
import pptx.pptxObject.PPTXFont;
import pptx.pptxObject.PPTXLatinFont;
import pptx.pptxObject.PPTXSchemeColorModel;
import pptx.pptxObject.PPTXShape;
import pptx.pptxObject.PPTXSlideMasterTextStyles;

/**
 * @author sheraz.ahmed
 *
 */
public class PPTXContext {
	
	private Thread conversionThread;
	private static Vector<PPTXContext> mOutputMappings = new Vector<PPTXContext>();
	private DriverOptions driverOptions = null;
	private String sourceFile = "";
	private HashMap<String, ByteArrayInputStream> pptxFilesStreams = new HashMap<String, ByteArrayInputStream>(); // stores stream for each unzipped file.	
	private ZipInputStream pptxZipStreamFile = null;
	private String destinationFile = "";
	private HashMap<String, ZipEntry> PPTXEntries = new HashMap<String, ZipEntry>();	//stores name, zip entry pair for all entries in docx file
	private File imageDirectory = null;
	private File m_imgSaveDir = null;	
	private ResourcePacket lPacket = null;	
	private Element m_ImgPacketDir = null;
	private PPXMLDocument ppXMLDocument;	//the destination document 
	private PPXMLObjectFactory factory; //used to create ppXML objects	
	private boolean isDefaultParagraphPropertiesPresent = false;
	private boolean isParagraphPropertiesPresent = false;
	private PPTXSlideMasterTextStyles PPTXSlideMasterTextStyles = null;
	private HashMap<String, String> pptxColors = new HashMap<String, String>();
	private HashMap<String, String> pptxMajorFonts = new HashMap<String, String>();
	private HashMap<String, String> pptxMinorFonts = new HashMap<String, String>();
	private HashMap<String, PPTXShape> pptxMasterSlideShapes = new HashMap<String, PPTXShape>();
	private HashMap<String, Document> slideLayouts = new HashMap<String, Document>();
	private List<String> imagesWrittinToDir = new ArrayList<String>();
	private int imageCounter = 1;
	private int rowCount = 0;
	private int cellCount = 0;
	private int tableHeight = 0;
	private List<String> columnWidths = new ArrayList<String>();
	private List<String> rowHeights = new ArrayList<String>();
	private int marker_IDSlideNo = -1;
	private int marker_LastSlideNo = -1;
	private int shapeNo = -1;	
	private int prevRowHeight = 0;
	private int prevCellY = 0;
	private int prevRowY = 0;
	private int prevColumnWidth = 0;
	private int prevCellX = 0;
	private boolean isInSameRow = true;
	/**
	 * Default Constructor.
	 * <p>
	 * Made private to make it available a single instance throughout the single iteration.
	 * To get this class object use the getInstance method 
	 * <p>
	 */
	private PPTXContext() 
	{
		resetContext();
	}
	
	/**
	 * 
	 * @param ctx
	 * @return
	 */
	private static boolean addOutputMapping(PPTXContext ctx)
	{
		mOutputMappings.add(ctx);
		return true;
	}
	
	/** Description: clear all the memory allocated by current context
	 * @author Shabana.Majeed
	 * 
	 */
	public void clearContext()
	{	
		// private Thread conversionThread;
		this.setConversionThread(null);
		
		// private static Vector<PPTXContext> mOutputMappings = new Vector<PPTXContext>();
		for (int i = 0; i < this.mOutputMappings.size(); i++)
		{
			PPTXContext rc = this.mOutputMappings.get(i);
			rc = null;			
		}
		
		// private String sourceFile = "";
		this.setSourceFile("");
		
		// private HashMap<String, ByteArrayInputStream> pptxFilesStreams = new HashMap<String, ByteArrayInputStream>();
		for (int i = 0; i < this.pptxFilesStreams.size(); i++)
		{
			ByteArrayInputStream byteArray = this.pptxFilesStreams.get(i);
			byteArray = null;			
		}		

		// private ZipInputStream pptxZipStreamFile = null;
		this.pptxZipStreamFile = null;
	
		// private String destinationFile = "";
		this.setDestinationFile("");
		
		// private HashMap<String, ZipEntry> PPTXEntries = new HashMap<String, ZipEntry>();
		for (int i = 0; i < this.PPTXEntries.size(); i++)
		{
			ZipEntry zipEntry = this.PPTXEntries.get(i);
			zipEntry = null;			
		}
		
		// private File imageDirectory = null;
		this.setImageDirectory(null);

		// private File m_imgSaveDir = null;
		this.m_imgSaveDir = null;
		
		// private ResourcePacket lPacket = null;
		this.setResourcePacket(null);
		
		// private Element m_ImgPacketDir = null;
		this.setM_ImgPacketDir(null);
		
		// private PPXMLDocument ppXMLDocument;
		this.ppXMLDocument = null;
		
		// private PPXMLObjectFactory factory;
		this.factory = null;

		// private boolean isDefaultParagraphPropertiesPresent = false;
		this.setDefaultParagraphPropertiesPresent(false);
		
		// private boolean isParagraphPropertiesPresent = false;
		this.setParagraphPropertiesPresent(false);
		
		// private PPTXSlideMasterTextStyles PPTXSlideMasterTextStyles = null;
		this.setPPTXSlideMasterTextStyles(null);
		
		// private HashMap<String, String> pptxColors = new HashMap<String, String>();
		for (int i = 0; i < this.pptxColors.size(); i++)
		{
			String pptxColor = this.pptxColors.get(i);
			pptxColor = null;			
		}
		
		// private HashMap<String, String> pptxMajorFonts = new HashMap<String, String>();
		for (int i = 0; i < this.pptxMajorFonts.size(); i++)
		{
			String pptxMajorFonts = this.pptxMajorFonts.get(i);
			pptxMajorFonts = null;			
		}
		
		// private HashMap<String, String> pptxMinorFonts = new HashMap<String, String>();
		for (int i = 0; i < this.pptxMinorFonts.size(); i++)
		{
			String pptxMinorFonts = this.pptxMinorFonts.get(i);
			pptxMinorFonts = null;			
		}
		
		// private HashMap<String, PPTXShape> pptxMasterSlideShapes = new HashMap<String, PPTXShape>();		
		for (int i = 0; i < this.pptxMasterSlideShapes.size(); i++)
		{
			PPTXShape masterSlideShape = this.pptxMasterSlideShapes.get(i);
			masterSlideShape = null;			
		}
		
		// private HashMap<String, Document> slideLayouts = new HashMap<String, Document>();
		for (int i = 0; i < this.slideLayouts.size(); i++)
		{
			Document slideLayout = this.slideLayouts.get(i);
			slideLayout = null;			
		}
		
		// private List<String> imagesWrittinToDir = new ArrayList<String>();
		for (int i = 0; i < this.imagesWrittinToDir.size(); i++)
		{
			String imageWrittenToDir = this.imagesWrittinToDir.get(i);
			imageWrittenToDir = null;			
		}
		
		// private int imageCounter = 1;
		this.setImageCounter(0);
		
		// private int rowCount = 0;
		this.setRowCount(0);
		
		// private int cellCount = 0;
		this.setCellCount(0);
		
		// private int tableHeight = 0;
		this.setTableHeight(0);
		
		// private List<String> columnWidths = new ArrayList<String>();
		for (int i = 0; i < this.columnWidths.size(); i++)
		{
			String colWidth = this.columnWidths.get(i);
			colWidth = null;			
		}
		
		// private List<String> rowHeights = new ArrayList<String>();
		for (int i = 0; i < this.rowHeights.size(); i++)
		{
			String rowHeight = this.rowHeights.get(i);
			rowHeight = null;			
		}
		
		// private int marker_IDSlideNo = -1;
		this.setMarker_IDSlideNo(0);
		
		// private int marker_LastSlideNo = -1;
		this.setMarker_LastSlideNo(0);
		
		// private int shapeNo = -1;
		this.setShapeNo(0);
		
		// private int prevRowHeight = 0;
		this.setPrevRowHeight(0);
		
		// private int prevCellY = 0;
		this.setPrevCellY(0);
		
		// private int prevRowY = 0;
		this.setPrevRowY(0);
		
		// private int prevColumnWidth = 0;
		this.setPrevColumnWidth(0);
		
		// private int prevCellX = 0;
		this.setPrevCellX(0);
		
		// private boolean isInSameRow = true;
		this.setInSameRow(false);
	}
	
	/**
	 * 
	 * @param conversionThread
	 */
	private static void removeOutputMapping(Thread param_conversionThread)
	{	
		PPTXContext rc = null;

		for (int i = 0; i < mOutputMappings.size(); i++)
		{
			rc = mOutputMappings.get(i);
			if (rc.getConversionThread() != null)
			{
				if(rc.getConversionThread().equals(param_conversionThread))	
				{
					rc.resetContext();
					mOutputMappings.remove(i);
					return;
				}
			}
		}		
	}
	
	/**
	 * 
	 * @param conversionThread
	 * @return
	 */
	private static PPTXContext getOutputMapping(Thread param_conversionThread)
	{	
		PPTXContext rc = null;
		PPTXContext result = null;

		for (int i = 0; i < mOutputMappings.size(); i++)
		{
			rc = mOutputMappings.get(i);
			if (rc.getConversionThread() != null)
			{
				if(rc.getConversionThread().equals(param_conversionThread))	
				{
					result = rc;
					break;
				}
			}				
		}

		return result;
	}	
	
	/**
	 * 
	 * @return
	 */
	public static PPTXContext getInstance()
	{
		Thread currThread = Thread.currentThread();
		PPTXContext context = getOutputMapping(currThread);


		if(context == null)
		{
			context = new PPTXContext();
			context.setConversionThread(currThread);
			addOutputMapping(context);
		}
		return context;
	}

	/**
	 * 
	 */
	public void resetContext()
	{
		this.ppXMLDocument = new PPXMLDocument();
		this.factory = ppXMLDocument.getObjectFactory();
		this.pptxColors = new HashMap<String, String>();
		this.pptxMajorFonts = new HashMap<String, String>();
		this.pptxMinorFonts = new HashMap<String, String>();
	}

	/** Description: reset all attributes of context that are being used in calculating cell margins
	 * <p>
     * @author Shabana Majeed
	 */
	public void resetTableCoordinatesData()
	{
		this.setCellCount(0);
		this.setRowCount(0);
		this.setPrevCellX(0);
		this.setPrevCellY(0);
		this.setPrevRowHeight(0);
		this.setPrevRowY(0);
		this.setColumnWidths(new ArrayList<String>());
		this.setRowHeights(new ArrayList<String>());		
	}
	
	/**
	 * 
	 */
	public static void removeInstance()
	{
		Thread currThread = Thread.currentThread();
		removeOutputMapping(currThread);
	}
	
	/**
	 * 
	 * @param fileStream of pptx file
	 * @return true if pptx file is unzipped, otherwise false
	 */
	public boolean unzipFileAndStoreStreams(InputStream fileStream)
	{
		ZipEntry currentEntry = null;
		this.pptxFilesStreams.clear();

		this.pptxZipStreamFile = new ZipInputStream(fileStream);
		
		try
		{			
			while(true)
			{
				currentEntry = this.pptxZipStreamFile.getNextEntry();
				if(currentEntry != null)
				{
					String name = currentEntry.getName();    
					this.PPTXEntries.put(name, currentEntry);

					long bufferSize = currentEntry.getSize();
					Long bufferSizeLong = new Long(bufferSize);
					int bufferSizeIntValue = bufferSizeLong.intValue();
					bufferSizeLong = null;
					
					byte[] buffer = new byte[bufferSizeIntValue]; 
					byte[] readBuffer = new byte[bufferSizeIntValue];
					
					int bytesRead = 0;
					while(bytesRead <  bufferSizeIntValue)
					{
						int currBytesRead = this.pptxZipStreamFile.read(readBuffer, 0 , bufferSizeIntValue);
						for(int index = 0 ; index < currBytesRead ; index++)
						{
							buffer[bytesRead] = readBuffer[index];
							bytesRead++;	    						
						}
					} 

					ByteArrayInputStream fileBytesStream = new ByteArrayInputStream(buffer);
					this.pptxFilesStreams.put(name, fileBytesStream);
				}
				else
				{
					break;
				}
			}

		}catch(Exception ex)
		{
			DEBUG.errorPrint("PPTX - Could not unzip pptx source file - " + ex.getMessage());
			return false;
		}
		return true;
	}
	
	/**
	 * 
	 * @param entry instance of ZipEntry
	 * @return
	 */
	public InputStream getXMLChunkStream(ZipEntry entry)
	{
		try
		{
			if(pptxFilesStreams != null)		
			{
				if(entry != null && entry.getName() != null && entry.getName().trim().length() > 0)
				{
					return (InputStream)pptxFilesStreams.get(entry.getName());
				}
			}
		}catch(Exception ex)
		{
			DEBUG.errorPrint("PPTX - Could not get zip entry: " + entry.getName()  + " Reason: " + ex.getMessage());
			return null;
		}
		return null;
	}
	
	
	/**
     * Description: Extracts the colors name with their hex values
     * then save the extracted colors in hash Map with their names as key
     * <p>
     * @param its input will be color schema node + its children from theme.xml
     * @param its output contains hash map having schema colors
     */

	
	public void loadPPTXColors(NodeList colorSchemeChildren)
	{
		for(int index = 0; index < colorSchemeChildren.getLength(); index++)
		{
			Node child = colorSchemeChildren.item(index);
			String nodeName = child.getNodeName();     //<a:dk1>
			Node colorChild = child.getFirstChild();  //a:sysClr or a:srgbClr
			
			NamedNodeMap colorAttributes = colorChild.getAttributes();
			
			if (colorAttributes != null)
			{
				if(colorChild.getNodeName().equalsIgnoreCase(PPTXConstants.PPTX_COLOR_SYSTEM))
				{					
					pptxColors.put(nodeName.substring(2),
													colorAttributes.getNamedItem("lastClr").getNodeValue());
					pptxColors.put(colorAttributes.getNamedItem("val").getNodeValue(),
													colorAttributes.getNamedItem("lastClr").getNodeValue());
				}
				
				else if(colorChild.getNodeName().equalsIgnoreCase(PPTXConstants.PPTX_COLOR_SYSTEMRGB))
				{
					pptxColors.put(nodeName.substring(2),colorAttributes.getNamedItem("val").getNodeValue());
				}
			}		
		}
	}
	
	/**
     * Description: Extracts the Major Fonts names with their script name and typeface
     * then save the extracted Major Fonts in hash Map with their scripts names as key
     * <p>
     * @param its input will be Major Font node + its children from theme.xml
     * @param its output contains hash map having schema Major Fonts
     */
	
	public void loadPPTXMajorFonts(NodeList majorFontChildren)
	{
		for(int index = 0; index < majorFontChildren.getLength(); index++)
		{
			Node child = majorFontChildren.item(index);
			String nodeName = child.getNodeName();
			NamedNodeMap childAttributes = child.getAttributes();
		
			if (childAttributes != null)
			{ 
				if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_LATIN))
				{
					pptxMajorFonts.put(PPTXConstants.PPTX_FONT_LATIN.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_EASTASSIAN))
				{
					pptxMajorFonts.put(PPTXConstants.PPTX_FONT_EASTASSIAN.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_COMPLEXSCRIPT))
				{
					pptxMajorFonts.put(PPTXConstants.PPTX_FONT_COMPLEXSCRIPT.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT))
				{
					pptxMajorFonts.put(childAttributes.getNamedItem("script").getNodeValue(),childAttributes.getNamedItem("typeface").getNodeValue());
				}
			}			
		}	
	}
	
	/**
     * Description: Extracts the Minor Fonts names with their script name and typeface
     * then save the extracted Minor Fonts in hash Map with their scripts names as key
     * <p>
     * @param its input will be Minor Font node + its children from theme.xml
     * @param its output contains hash map having schema Minor Fonts
     */
	
	public void loadPPTXMinorFonts(NodeList minorFontChildren)
	{
		for(int index = 0; index < minorFontChildren.getLength(); index++)
		{
			Node child = minorFontChildren.item(index);
			String nodeName = child.getNodeName();
			NamedNodeMap childAttributes = child.getAttributes();
			if (childAttributes != null)
			{
				if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_LATIN))
				{
					pptxMinorFonts.put(PPTXConstants.PPTX_FONT_LATIN.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_EASTASSIAN))
				{
					pptxMinorFonts.put(PPTXConstants.PPTX_FONT_EASTASSIAN.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT_COMPLEXSCRIPT))
				{
					pptxMinorFonts.put(PPTXConstants.PPTX_FONT_COMPLEXSCRIPT.substring(2),childAttributes.getNamedItem("typeface").getNodeValue());
				}
				else if(nodeName.equalsIgnoreCase(PPTXConstants.PPTX_FONT))
				{
					pptxMinorFonts.put(childAttributes.getNamedItem("script").getNodeValue(),childAttributes.getNamedItem("typeface").getNodeValue());
				}
			}
		}		
	}
	
	public String getMajorFontTypeFace(String fontName)
	{
		String fontTypeFace = "";
		if(this.pptxMajorFonts.containsKey(fontName))
		{
			fontTypeFace = this.pptxMajorFonts.get(fontName);
		}
		return fontTypeFace;
	}
	
	public String getMinorFontTypeFace(String fontName)
	{
		String fontTypeFace = "";
		if(this.pptxMinorFonts.containsKey(fontName))
		{
			fontTypeFace = this.pptxMinorFonts.get(fontName);
		}
		return fontTypeFace;
	}	
	
	public String getColorValue(String colorName)
	{
		String colorValue = "";
		if(colorName.equals("bg1"))
		{
			colorName = "lt1";
		}
		if(this.pptxColors.containsKey(colorName))
		{
			colorValue = this.pptxColors.get(colorName);
		}
		return colorValue;
	}


	/**
	 * ------------------------------------------------------------------------------------
	 * 								All getter and setter methods
	 * ------------------------------------------------------------------------------------
	 */
	
	/**
	 * 
	 * @param options
	 */
	public void setDriverOptions(String options)
	{
		this.driverOptions = new DriverOptions(options);
	}
	
	/**
	 * 
	 * @return sourceFile
	 */
	public String getSourceFile() {
		return sourceFile;
	}

	/**
	 * 
	 * @param sourceFile
	 */
	public void setSourceFile(String sourceFile) {
		this.sourceFile = sourceFile;
	}

	/**
	 * 
	 * @return destinationFile
	 */
	public String getDestinationFile() {
		return destinationFile;
	}

	/**
	 * 
	 * @param destinationFile
	 */
	public void setDestinationFile(String destinationFile) {
		this.destinationFile = destinationFile;
	}

	/**
	 * 
	 * @return DriverOptions object
	 */
	public DriverOptions getDriverOptions() {
		return driverOptions;
	}

	/**
	 * 
	 * @return PPTXEntries object
	 */
	public HashMap<String, ZipEntry> getPPTXEntries() {
		return PPTXEntries;
	}

	/**
	 * 
	 * @return File instance of image Saving Directory
	 */
    public File getImageSaveDir()
	{
		return this.m_imgSaveDir;
	}
	
    /**
     * 
     * @param File instance of image Saving Directory
     */
	public void setImageSaveDir(File imageSaveDir)
	{
		this.m_imgSaveDir = imageSaveDir;
	}

	/**
	 * 
	 * @param File instance of image Directory
	 */
	public void setImageDirectory(File imageDirectory) {
		this.imageDirectory = imageDirectory;
	}

	/**
	 * 
	 * @param lpacket
	 */
	public void setResourcePacket(ResourcePacket lpacket)
	{
		this.lPacket = lpacket;
	}
	
	/**
	 * 
	 * @return lPacket instance
	 */
	public ResourcePacket getResourcePacket()
	{
		return this.lPacket;
	}
	
	/**
	 * 
	 * @return
	 */
	public Element getM_ImgPacketDir() 
	{
		return m_ImgPacketDir;
	}

	/**
	 * 
	 * @param m_ImgPacketDir
	 */
	public void setM_ImgPacketDir(Element m_ImgPacketDir) 
	{
		this.m_ImgPacketDir = m_ImgPacketDir;
	}

	/**
	 * 
	 * @return
	 */
	public Thread getConversionThread() 
	{
		return conversionThread;
	}
	
	/**
	 * 
	 * @param conversionThread
	 */
	public void setConversionThread(Thread param_conversionThread) 
	{
		this.conversionThread = param_conversionThread;
	}

	/**
	 * 
	 * @return PPXMLObjectFactory instance
	 */
	public PPXMLObjectFactory getFactory() 
	{
		return factory;
	}

	/**
	 * 
	 * @return ppXMLDocument
	 */
	public PPXMLDocument getPPXMLDocument() {
		return ppXMLDocument;
	}

	/**
	 * 
	 * @return isDefaultParagraphPropertiesPresent
	 */
	public boolean isDefaultParagraphPropertiesPresent() 
	{
		return isDefaultParagraphPropertiesPresent;
	}

	/**
	 * 
	 * @param isDefaultParagraphPropertiesPresent
	 */
	public void setDefaultParagraphPropertiesPresent(
			boolean isDefaultParagraphPropertiesPresent) 
	{
		this.isDefaultParagraphPropertiesPresent = isDefaultParagraphPropertiesPresent;
	}

	/**
	 * 
	 * @return isParagraphPropertiesPresent
	 */
	public boolean isParagraphPropertiesPresent() 
	{
		return isParagraphPropertiesPresent;
	}

	/**
	 * 
	 * @param isParagraphPropertiesPresent
	 */
	public void setParagraphPropertiesPresent(boolean isParagraphPropertiesPresent) 
	{
		this.isParagraphPropertiesPresent = isParagraphPropertiesPresent;
	}
	
	/**
	 * @return the pPTXSlideMasterTextStyles
	 */
	public PPTXSlideMasterTextStyles getPPTXSlideMasterTextStyles() 
	{
		return PPTXSlideMasterTextStyles;
	}

	/**
	 * @param pPTXSlideMasterTextStyles the pPTXSlideMasterTextStyles to set
	 */
	public void setPPTXSlideMasterTextStyles(
			PPTXSlideMasterTextStyles pPTXSlideMasterTextStyles) 
	{
		this.PPTXSlideMasterTextStyles = pPTXSlideMasterTextStyles;
	}

	/**
	 * @return the pptxMasterSlideShapes
	 */
	public HashMap<String, PPTXShape> getPptxMasterSlideShapes() 
	{
		return pptxMasterSlideShapes;
	}

	/**
	 * @param pptxMasterSlideShapes the pptxMasterSlideShapes to set
	 */
	public void setPptxMasterSlideShapes(
			HashMap<String, PPTXShape> pptxMasterSlideShapes) 
	{
		this.pptxMasterSlideShapes = pptxMasterSlideShapes;
	}
	
	
	/**
	 * @return the slideLayouts
	 */
	public HashMap<String, Document> getSlideLayouts() {
		return slideLayouts;
	}

	/**
	 * @param slideLayouts the slideLayouts to set
	 */
	public void setSlideLayouts(HashMap<String, Document> slideLayouts) {
		this.slideLayouts = slideLayouts;
	}
	
	/**
	 * @return the imagesWrittinToDir
	 */
	public List<String> getImagesWrittinToDir() {
		return imagesWrittinToDir;
	}

	/**
	 * @param imagesWrittinToDir the imagesWrittinToDir to set
	 */
	public void setImagesWrittinToDir(List<String> imagesWrittinToDir) {
		this.imagesWrittinToDir = imagesWrittinToDir;
	}

	
	
	/**
	 * @return the imageCounter
	 */
	public int getImageCounter() {
		return imageCounter;
	}

	/**
	 * @return Returns the cellCount.
	 */
	public int getCellCount() {
		return cellCount;
	}

	/**
	 * @param cellCount The cellCount to set.
	 */
	public void setCellCount(int cellCount) {
		this.cellCount = cellCount;
	}

	/**
	 * @return Returns the rowCount.
	 */
	public int getRowCount() {
		return rowCount;
	}

	/**
	 * @param rowCount The rowCount to set.
	 */
	public void setRowCount(int rowCount) {
		this.rowCount = rowCount;
	}

	/**
	 * @param imageCounter the imageCounter to set
	 */
	public void setImageCounter(int imageCounter) {
		this.imageCounter = imageCounter;
	}

	/**
	 * @return Returns the tableHeight.
	 */
	public int getTableHeight() {
		return tableHeight;
	}

	/**
	 * @param tableHeight The tableHeight to set.
	 */
	public void setTableHeight(int tableHeight) {
		this.tableHeight = tableHeight;
	}

	/**
	 * @return Returns the marker_IDSlideNo.
	 */
	public int getMarker_IDSlideNo() {
		return marker_IDSlideNo;
	}

	/**
	 * @param marker_IDSlideNo The marker_IDSlideNo to set.
	 */
	public void setMarker_IDSlideNo(int marker_IDSlideNo) {
		this.marker_IDSlideNo = marker_IDSlideNo;
	}

	/**
	 * @return Returns the shapeNo.
	 */
	public int getShapeNo() {
		return shapeNo;
	}

	/**
	 * @param shapeNo The shapeNo to set.
	 */
	public void setShapeNo(int shapeNo) {
		this.shapeNo = shapeNo;
	}

		
	/**
	 * @return Returns the columnWidths.
	 */
	public List<String> getColumnWidths() {
		return columnWidths;
	}

	/**
	 * @param columnWidths The columnWidths to set.
	 */
	public void setColumnWidths(List<String> columnWidths) {
		this.columnWidths = columnWidths;
	}

	/**
	 * @return Returns the rowHeights.
	 */
	public List<String> getRowHeights() {
		return rowHeights;
	}

	/**
	 * @param rowHeights The rowHeights to set.
	 */
	public void setRowHeights(List<String> rowHeights) {
		this.rowHeights = rowHeights;
	}

	/**
	 * @return Returns the prevCellX.
	 */
	public int getPrevCellX() {
		return prevCellX;
	}

	/**
	 * @param prevCellX The prevCellX to set.
	 */
	public void setPrevCellX(int prevCellX) {
		this.prevCellX = prevCellX;
	}

	/**
	 * @return Returns the prevCellY.
	 */
	public int getPrevCellY() {
		return prevCellY;
	}

	/**
	 * @param prevCellY The prevCellY to set.
	 */
	public void setPrevCellY(int prevCellY) {
		this.prevCellY = prevCellY;
	}

	/**
	 * @return Returns the prevColumnWidth.
	 */
	public int getPrevColumnWidth() {
		return prevColumnWidth;
	}

	/**
	 * @param prevColumnWidth The prevColumnWidth to set.
	 */
	public void setPrevColumnWidth(int prevColumnWidth) {
		this.prevColumnWidth = prevColumnWidth;
	}

	/**
	 * @return Returns the prevRowHeight.
	 */
	public int getPrevRowHeight() {
		return prevRowHeight;
	}

	/**
	 * @param prevRowHeight The prevRowHeight to set.
	 */
	public void setPrevRowHeight(int prevRowHeight) {
		this.prevRowHeight = prevRowHeight;
	}

	/**
	 * @return Returns the isInSameRow.
	 */
	public boolean isInSameRow() {
		return isInSameRow;
	}

	/**
	 * @param isInSameRow The isInSameRow to set.
	 */
	public void setInSameRow(boolean isInSameRow) {
		this.isInSameRow = isInSameRow;
	}

	/**
	 * @return Returns the prevRowY.
	 */
	public int getPrevRowY() {
		return prevRowY;
	}

	/**
	 * @return Returns the marker_LastSlideNo.
	 */
	public int getMarker_LastSlideNo() {
		return marker_LastSlideNo;
	}

	/**
	 * @param marker_LastSlideNo The marker_LastSlideNo to set.
	 */
	public void setMarker_LastSlideNo(int marker_LastSlideNo) {
		this.marker_LastSlideNo = marker_LastSlideNo;
	}

	/**
	 * @param prevRowY The prevRowY to set.
	 */
	public void setPrevRowY(int prevRowY) {
		this.prevRowY = prevRowY;
	}

	public PPTXShape getMasterSlideShapeByType(String type)
	{
		if(type == null && type.trim().length() == 0)
			return null;
		
		PPTXShape pptxShape = null;
		Iterator<String> iterator = this.pptxMasterSlideShapes.keySet().iterator();
		while(iterator.hasNext())
		{
			String key = iterator.next();
			if(type.toLowerCase().equals("subTitle"))
			{
				type = "body";
			}
			if(type.toLowerCase().contains(key.toLowerCase()))
			{
				pptxShape = this.pptxMasterSlideShapes.get(key);
				break;
			}
		}
		return pptxShape;
	}

	
}
