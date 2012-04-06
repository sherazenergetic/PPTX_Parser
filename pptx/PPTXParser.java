/**
 * @author Sheraz Ahmed
 */
package pptx;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.zip.ZipEntry;

import org.w3c.dom.Document;

import pptx.PPTXConstants;
import pptx.parser.CorePropertiesParser;
import pptx.context.PPTXContext;
import pptx.exception.BlockerException;
import pptx.parser.PPTXParserFactory;
import pptx.parser.Parser;
import pptx.parser.ParserInputData;
import pptx.pptxObject.PPTXSlideMasterTextStyles;

/**
 * @author sheraz.ahmed
 *
 */
public class PPTXParser 
{
	private XYZInputSource outputPPXML;

	private PPTXContext context = PPTXContext.getInstance();
	private PPTXParserFactory parserFactory = PPTXParserFactory.getInstance();
	
	/**
	 * Default Constructor
	 */
	public PPTXParser()
	{
		
	}
	/**
	 * The method parses the pptx file.
	 * <p>
	 * It will do the following steps:<p>
	 * <li>Unzips the file and records all entries in a hash map</li>
	 * <li>Creates PPXML HEAD element on the basis of core and app files from pptx</li>
	 * <li>Creates PPXML BODY element on the basis of document file from the pptx</li>
	 * <li>Attaches them in the PPXMLDocument and writes it to output</li>
	 * <p>
	 * <p>
	 * @param input		input file, e.g. any pptx file.
	 * @param output	output file, e.g. xml file.
	 * @param options	driver options
	 * <p>
	 * @return 			boolean result whether the parsing completed or not.
	 */

	public boolean parse(String input, String output, String options) 
	{
		return parse(new XYZInputSource(input), new XYZInputSource(output), options);
	}

	/**
	 * The method parses the pptx file.
	 * <p>
	 * It will do the following steps:<p>
	 * <li>Unzips the file and records all entries in a hash map</li>
	 * <li>Creates PPXML HEAD element on the basis of core and app files from pptx</li>
	 * <li>Creates PPXML BODY element on the basis of document file from the pptx</li>
	 * <li>Attaches them in the PPXMLDocument and writes it to output</li>
	 * <p>
	 * <p>
	 * @param input		input XYZInputSource object containing the pptx file
	 * @param output	input XYZInputSource object containing the pptx file
	 * @param options	driver options
	 * <p>
	 * @return 			boolean result whether the parsing completed or not.
	 */
	public boolean parse(XYZInputSource input, XYZInputSource output, String options)
	{
		outputPPXML = output;
		try
		{
	    	boolean contextCreated = buildContext(input, output, options);
	    	if(!contextCreated)
	    	{
	    		return false;
	    	}

	    	String lImageDirParam = context.getDriverOptions().IMAGE_DIRECTORY;	
	    	setImageDir(lImageDirParam);
	    	
	    	HashMap<String, ZipEntry> pptxEntries = context.getPPTXEntries();	
	    	ZipEntry relsZipEntry = pptxEntries.get("_rels/.rels");
	    	
	    	HashMap<String, String> globalRelationships = reteriveGlobalRelationships(relsZipEntry);
	    	
			PPXMLObject metaTags = null;
			
			String target = (String) globalRelationships.get(PPTXConstants.PPTX_RELATIONSHIPS_CORE);
			
			if (target != null)
			{				
				ZipEntry coreZipEntry = (ZipEntry)pptxEntries.get(target);
				metaTags = loadCoreRelationshipFileMeta(coreZipEntry);
			}
			
			target = (String) globalRelationships.get(PPTXConstants.PPTX_RELATIONSHIPS_EXTENDED);

			if (target != null)
			{				
				ZipEntry appZipEntry = (ZipEntry)pptxEntries.get(target);
				PPXMLObject appExtendedMetaTags = loadAppExtendedRelationshipFileMeta(appZipEntry);
				if(metaTags != null && appExtendedMetaTags != null)
				{
					metaTags.copyChildren(appExtendedMetaTags);
				}
			}
			
			
	    	ZipEntry masterSlideRelsZipEntry = pptxEntries.get("ppt/slideMasters/_rels/slideMaster1.xml.rels");
	    	
	    	HashMap<String, String> slideMaster1Relationships = reteriveGlobalRelationships(masterSlideRelsZipEntry);
			Iterator<String> iterator = slideMaster1Relationships.keySet().iterator();
			String masterSlideThemeEntryName = "";
			while(iterator.hasNext())
			{
				String key = iterator.next();
				if(key.contains(PPTXConstants.PPTX_RELATIONSHIPS_THEME))
				{
					masterSlideThemeEntryName = slideMaster1Relationships.get(key);
					break;
				}
			}
			masterSlideThemeEntryName = masterSlideThemeEntryName.substring(masterSlideThemeEntryName.indexOf("/"));
			masterSlideThemeEntryName = "ppt" + masterSlideThemeEntryName;
			ZipEntry masterSlideThemeZipEntry = pptxEntries.get(masterSlideThemeEntryName);
			loadPPTXSlideMasterTheme(masterSlideThemeZipEntry);

			ZipEntry slideMasterZipEntry = pptxEntries.get("ppt/slideMasters/slideMaster1.xml");
			PPXMLObject pptxSlideMasterTextStyles = null;
			if(slideMasterZipEntry != null)
			{
				pptxSlideMasterTextStyles = loadPPTXSlideMasterTextStyles(slideMasterZipEntry);
			}
	 
			
			
//			Enumeration<?> enu = metaTags.getChildren();
//			while(enu.hasMoreElements())
//			{
//				PPXMLMeta object = (PPXMLMeta)enu.nextElement();
//				System.out.println(object.getProperty(ppXML.PPXML_ATT_NAME) +"="+ object.getProperty(ppXML.PPXML_ATT_VALUE)); 
//			}
			
			target = (String) globalRelationships.get(PPTXConstants.PPTX_RELATIONSHIPS_DOCUMENT);
			ZipEntry presentationEntry = (ZipEntry)pptxEntries.get(target);
			
			ZipEntry presentation_relsZipEntry = pptxEntries.get("ppt/_rels/presentation.xml.rels");
	    	
			//NOTE: need to manage slide information in List instead of HashMap. It is not consistent.
	    	HashMap<String, String> presentationRelationships = reteriveGlobalRelationships(presentation_relsZipEntry); 

	    	//process head
			PPXMLObjectFactory factory = context.getFactory();
			PPXMLObject head = factory.createObject(ppXML.PPXML_HEAD);
	    	
	    	PPXMLObject body = null;
	    	
	    	if(target != null && presentationRelationships != null)
	    	{
	    		body = parsePresentationEntry(presentationEntry, presentationRelationships);
	    	}
	    	
			if (metaTags != null)
			{
				head.addChild(metaTags);
			}			
			head.addChild(getXYZMetaTags(input.getFilename(),options));
			
			if(pptxSlideMasterTextStyles != null)
			{
				head.addChild(pptxSlideMasterTextStyles);
			}
			
			PPXMLDocument ppxmlDocument = context.getPPXMLDocument();			
			ppxmlDocument.addPPXMLObj(head);
			ppxmlDocument.addPPXMLObj(body);

			writeToOutput();
			
			if(context.getResourcePacket() != null)
			{
				context.getResourcePacket().saveToFile();
			}
			
			context.clearContext();
			context.removeInstance();
    	}
		catch(BlockerException ex)
		{
			DEBUG.errorPrint("Error - PPTX Driver - An error occured while processing the pptx source. \n " + ex.getMessage());
			DEBUG.printException(ex);
			ex.printStackTrace();
			return false;
		}
		catch(Exception e)
		{
			DEBUG.errorPrint("Error - PPTX Driver - An error occured while processing the pptx source.");
			DEBUG.printException(e);
			return false;
		}

		return true;
	}
	
	/**
	 * 
	 * @param input
	 * @param output
	 * @param options
	 * @return
	 * @throws Exception
	 */
	private boolean buildContext(XYZInputSource input, XYZInputSource output, String options) throws Exception
	{
		InputStream sourceFileStream = null;
    	
    	switch(input.fileType)
    	{
    		case XYZInputSource.TYPE_FILENAME:
    			sourceFileStream = new FileInputStream(input.getFilename());    			
    			break;	    			 
    		case XYZInputSource.TYPE_INPUTSTREAM:
    			sourceFileStream = input.myInputStream;
    			break;
    		default:
    		{
    			DEBUG.errorPrint("PPTX: Unsupported input type: " + input.getTypeString());
    			return false;
    		}
    	}
		
		context.setDriverOptions(options);
		
    	if(input.myFileName != null)
		{
    		context.setSourceFile(input.myFileName);	    		    		
    	}	    	

    	if(output.getFilename()!=null)
    	{
    		context.setDestinationFile(output.myFileName);
    	}
    	
    	if(sourceFileStream != null)
    	{
    		boolean isFileUnzipped = context.unzipFileAndStoreStreams(sourceFileStream);
    		if(!isFileUnzipped)
    		{
    			return false;
    		}
    	}
    	else 
    	{
    		DEBUG.errorPrint("PPTX: There was a problem while unzip.");
    		return false;
		}
		return true;
	}
	
	/**
	 * 
	 */
	public void finalize()
	{
		PPTXContext.removeInstance();
	}
	
	/**
	 * 
	 * @param optionsImagDir
	 */
	private void setImageDir(String optionsImagDir)
	{
		// Set save path
		File file = new File(context.getSourceFile());

		File m_imgSaveDir = null;
		String m_imgSaveDirStr = "";
		if(context.getDestinationFile() != null && !context.getDestinationFile().equals(""))
		{
			m_imgSaveDir = new File(XYZUtilities.getParentDirectory(context.getDestinationFile()));
			m_imgSaveDirStr = XYZUtilities.getParentDirectory(context.getDestinationFile());
		}

		// If empty image dir string then use default
		if (optionsImagDir.equals(""))
		{
			String defImageDir = file.getName() + "_Images";
			m_imgSaveDir = new File(m_imgSaveDirStr + File.separator + defImageDir);
			m_imgSaveDirStr = defImageDir;
		}
		else if (optionsImagDir.equals(".")) // MM 27MAR06  image Naming Conventions
		{
			m_imgSaveDirStr = optionsImagDir;
			m_imgSaveDir = new File(XYZUtilities.getParentDirectory(context.getDestinationFile()));			 
		}
		else if (optionsImagDir.startsWith(".")) // MM 27MAR06  image Naming Conventions
		{
			m_imgSaveDirStr = optionsImagDir;
			File parentDirHelper = new File(XYZUtilities.getParentDirectory(context.getDestinationFile()));
			m_imgSaveDir = new File(parentDirHelper.getParent() );
		}
		else
		{
			if(!context.getDriverOptions().IMAGE_ABSOLUTE_PATHS) 
			{
				DEBUG.errorPrint(" When image directory is specified in \"Image Directory\" option, \"Set Absolute Path Images\" option cannot be false -- setting to true ");
				context.getDriverOptions().IMAGE_ABSOLUTE_PATHS = true;
			}
			m_imgSaveDirStr = context.getDriverOptions().IMAGE_DIRECTORY;

			// Check if image save dir is absolute or relative
			if (m_imgSaveDirStr.length() > 2)
			{
				if ((m_imgSaveDirStr.charAt(1) == ':') ||
					(m_imgSaveDirStr.startsWith("/")) ||
					(m_imgSaveDirStr.startsWith("\\")))
				{
					m_imgSaveDir = new File(m_imgSaveDirStr);
				}
				else
				{
					m_imgSaveDir = new File(file.getParent() + File.separator + m_imgSaveDirStr);
				}
			}
			else
			{
				m_imgSaveDir = new File(file.getParent() + File.separator + m_imgSaveDirStr);
			}
		}
		context.setImageDirectory(m_imgSaveDir);
		context.setImageSaveDir(m_imgSaveDir);
		
		if(!m_imgSaveDir.exists())
		{
			m_imgSaveDir.mkdirs();
		}
		
		
		if (context.getResourcePacket() != null)
		{
			context.setM_ImgPacketDir(context.getResourcePacket().addResourceDirectory(m_imgSaveDir.getAbsolutePath()));
		}

	}

	private HashMap<String, String> reteriveGlobalRelationships(ZipEntry relsZipEntry) throws BlockerException
	{
    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_RELATIONSHIPSPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setRelsZipEntry(relsZipEntry);
    	parser.populateData(parserInputData);
    	
    	boolean isParsingComplete = parser.parse();
    	
    	HashMap<String, String> globalRelationships = null;
    	if(isParsingComplete)
    	{
    		globalRelationships = parser.getOutputData().getMainRelationShips();
    	}
		
    	return globalRelationships;
	}
	
	private PPXMLObject loadCoreRelationshipFileMeta(ZipEntry coreZipEntry) throws BlockerException
	{
		PPXMLObject metaTags = null;

    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_COREPROPERTIESPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setCoreZipEntry(coreZipEntry);
    	parser.populateData(parserInputData);

    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		metaTags = parser.getOutputData().getMetaTags();
    	}

    	return metaTags;
	}

	private PPXMLObject loadAppExtendedRelationshipFileMeta(ZipEntry appZipEntry) throws BlockerException
	{
		PPXMLObject metaTags = null;

    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_APPEXTENDED_PROPERTIESPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setAppZipEntry(appZipEntry);
    	parser.populateData(parserInputData);

    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		metaTags = parser.getOutputData().getMetaTags();
    	}

    	return metaTags;
	}

	private PPXMLObject parsePresentationEntry(ZipEntry presentationEntry, HashMap<String, String> presentationRelationships) throws BlockerException
	{
		PPXMLObject body = null;

    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_PRESENTATIONPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setPresentationEntry(presentationEntry);
    	parserInputData.setPresentationRelationships(presentationRelationships);
    	parser.populateData(parserInputData);

    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		body = parser.getOutputData().getPPXMLBody();
    	}

    	return body;
	}
	
    private PPXMLObject getXYZMetaTags(String absolutepPathOfFile, String options)
    {    	
    	//creating rels target string dynamically    	 
    	int index = absolutepPathOfFile.lastIndexOf(File.separatorChar);
    	if(index == -1)
    	{
    		index = absolutepPathOfFile.lastIndexOf('/');
    	}
    	String directory = absolutepPathOfFile.substring(0, index+1);
    	String fileName = absolutepPathOfFile.substring(index+1);
    	PPXMLObjectFactory factory = context.getFactory();
    	PPXMLObject metaTags = factory.createObject(ppXML.PPXML_XYZ_METATAGS);
    	
    	try
    	{
    		PPXMLXYZMeta metaElement = (PPXMLXYZMeta)factory.createObject(ppXML.PPXML_XYZMETA);				
			//metaElement.setMeta(ppXML.PPXML_XYZMETA_STDPREPROCESSOR, ppXML.PPXML_DOCX_DRIVER);				
    		//metaElement.setMeta(ppXML.PPXML_XYZMETA_STDPREPROCESSOR, ppXML.PPXML_WORD_DRIVER);
    		metaElement.setMeta(ppXML.PPXML_XYZMETA_STDPREPROCESSOR, ppXML.PPXML_PPTX_DRIVER);
			metaTags.addChild(metaElement);
			
			metaElement = (PPXMLXYZMeta)factory.createObject(ppXML.PPXML_XYZMETA);				
			metaElement.setMeta(ppXML.PPXML_XYZMETA_STDSOURCEFILENAME, fileName);				
			metaTags.addChild(metaElement);
			
			metaElement = (PPXMLXYZMeta)factory.createObject(ppXML.PPXML_XYZMETA);				
			metaElement.setMeta(ppXML.PPXML_XYZMETA_STDSOURCEFILEDIR, directory);				
			metaTags.addChild(metaElement);
			
			metaElement = (PPXMLXYZMeta)factory.createObject(ppXML.PPXML_XYZMETA);				
			metaElement.setMeta(ppXML.PPXML_XYZMETA_STDPRE_TIME, DateFormat.getDateTimeInstance().format(new Date()));
			metaTags.addChild(metaElement);
			
			metaElement = (PPXMLXYZMeta)factory.createObject(ppXML.PPXML_XYZMETA);				
			metaElement.setMeta(ppXML.PPXML_XYZMETA_STDPRE_OPTIONS, options);				
			metaTags.addChild(metaElement);
    	}
    	catch(Exception e)
    	{
    		DEBUG.warningPrint("Warnning - PPTX - Could not set one or more Meta information elements");
    		DEBUG.printException(e);
    	}
    	return metaTags;
    }
    /**
     * Write the ppxml file to disk or stream or writer or dom depending
     * on the type specified by output XYZInputStream object
     *  
     */
    public void writeToOutput()
    {
    	PPXMLDocument ppXMLDocument = context.getPPXMLDocument();
    	ppXMLDocument.getPPXMLOptions().setOption(PPXMLOptions.kIncTab, "true");
    	ppXMLDocument.setOutput(outputPPXML.getFilename());
        Document dom = ppXMLDocument.getDOM();
        
        try
        {
	        switch(outputPPXML.fileType)
			{
				case XYZInputSource.TYPE_OUTPUTSTREAM:
					XYZDOMUtilities.saveDOMToOutputStream(dom, outputPPXML.myOutputStream);
					break;
				case XYZInputSource.TYPE_FILENAME:
					XYZDOMUtilities.saveDOMToFile(dom, new File(outputPPXML.getFilename()));
					break;
				case XYZInputSource.TYPE_WRITER:
					XYZDOMUtilities.saveDOMToWriter(dom, outputPPXML.myWriter);
					break;
				case XYZInputSource.TYPE_DOM:
					outputPPXML.myDoc = dom;
					break;
				default:
					DEBUG.errorPrint("PPTX: Unsupported output type: " + outputPPXML.getTypeString());	
			}
        }
        catch(IOException e)
        {
        	DEBUG.errorPrint("Error - PPTX - Error while writing dom to output.");
        }
    }
    
    private PPXMLObject loadPPTXSlideMasterTextStyles(ZipEntry slideMasterZipEntry) throws BlockerException
    {
    	PPXMLObject ppxmlObject = null;
    	
    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_SLIDEMASTER_TEXTSTYLESPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setSlideMasterZipEntry(slideMasterZipEntry);
    	parser.populateData(parserInputData);
    	
    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		ppxmlObject = parser.getOutputData().getPPTXSlideMasterTextStyles();
    	}
    	
    	return ppxmlObject;
    }

    private boolean loadPPTXSlideMasterTheme(ZipEntry masterSlideThemeZipEntry) throws BlockerException
    {
    	
    	Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_SLIDEMASTER_THEMEPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setMasterSlideThemeZipEntry(masterSlideThemeZipEntry);
    	parser.populateData(parserInputData);
    	
    	return parser.parse();
    }

}
