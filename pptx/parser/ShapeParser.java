/**
 * 
 */
package pptx.parser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import pptx.PPTXConstants;
import pptx.PPTXUtils;
import pptx.context.PPTXContext;
import pptx.exception.BlockerException;
import pptx.pptxObject.PPTXDefaultParagraphStyle;
import pptx.pptxObject.PPTXPicture;
import pptx.pptxObject.PPTXShape;
import pptx.pptxObject.PPTXTextParagraph;
import pptx.pptxObject.PPTXTextRun;

/**
 * @author sheraz.ahmed
 *
 */
public class ShapeParser implements Parser {

	private PPTXContext context = PPTXContext.getInstance();
	private PPXMLObjectFactory ppXMLObjectFactory = context.getFactory();
	private PPXMLShape shape = null;
	private PPTXShape pptxShape = null;
	private HashMap<String, String> SlideRelationships = null;
	private PPTXShape pptxLayoutShape = null;
	private PPTXParserFactory parserFactory = PPTXParserFactory.getInstance();
	
	/** 
	 * @see pptx.parser.Parser#populateData(pptx.parser.ParserInputData)
	 */
	public void populateData(ParserInputData parserInputData)
			throws BlockerException 
	{
		if(parserInputData != null)
		{
			if(parserInputData.getShapeElement() != null)
			{
				this.pptxShape = parserInputData.getShapeElement(); 
			}
			if(parserInputData.getShapeElementOfLayout() != null)
			{
				this.pptxLayoutShape = parserInputData.getShapeElementOfLayout(); 
			}
			if(parserInputData.getSlideRelationships() != null)
			{
				this.SlideRelationships = parserInputData.getSlideRelationships(); 
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
			shape = (PPXMLShape) ppXMLObjectFactory.createObject(ppXML.PPXML_SHAPE);
            shape.setProperty("vertical-relative", "page");
		 
          /**
		   * Change started by <Shabana Majeed> on  <05/27/2011>
	       * Reason : <Add support for colored background of shape>
		   * fix for Issue # 24 --- Fill colors for footers are not supported.
		   */                       
            
            if(!pptxShape.getShapeProperties().getSolidFill().getPptxRGBColorModelHexVariant().getValue().equals("") 
            		&&
            		(pptxShape.getShapeProperties().getSolidFill().getPptxRGBColorModelHexVariant().getValue() != null))
            {
            	shape.setProperty(ppXML.PPXML_ATT_BACKCOLOR,
            			pptxShape.getShapeProperties().getSolidFill().getPptxRGBColorModelHexVariant().getValue());
            }
            
            if(!pptxShape.getShapeProperties().getSolidFill().getPptxSchemeColorModel().getValue().equals("") 
            		&&
            		(pptxShape.getShapeProperties().getSolidFill().getPptxSchemeColorModel().getValue() != null))
            {
             	shape.setProperty(ppXML.PPXML_ATT_BACKCOLOR,
            			pptxShape.getShapeProperties().getSolidFill().getPptxSchemeColorModel().getValue());
            }
            
            if(!pptxShape.getShapeProperties().getSolidFill().getPptxSystemColorModel().getValue().equals("") 
            		&&
            		(pptxShape.getShapeProperties().getSolidFill().getPptxSystemColorModel().getValue() != null))
            {
            	shape.setProperty(ppXML.PPXML_ATT_BACKCOLOR,
            			pptxShape.getShapeProperties().getSolidFill().getPptxSystemColorModel().getValue());
            }
            
		    /**
		     * Change ended by <Shabana Majeed> on  <05/27/2011>
		     */
            
            if(pptxShape != null && pptxShape instanceof PPTXPicture)
			{
				int x = 0;
				int y = 0;
				int height = 0;
				int width = 0;
				String type = this.pptxShape.getPptxNonVisualProperties().getPptxPlaceholderShape().getType();

				PPTXShape masterSlidePPTXShape = context.getMasterSlideShapeByType(type);
				x = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
				if(x < 0)
				{
					x = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
					if(x < 0)
					{
						x = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
					}
				}
				y = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
				if(y < 0)
				{
					y = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
					if(y < 0)
					{
						y = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
					}
				}
				
				width = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
				if(width < 0)
				{
					width = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
					if(width < 0)
					{
						width = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
					}
				}
				
				height = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
				if(height < 0)
				{
					height = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
					if(height < 0)
					{
						height= Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
					}
				}
				
				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_HEIGHT, PPTXUtils.convertEMUsToPoints(height));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_HEIGHT);
				}

				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_WIDTH, PPTXUtils.convertEMUsToPoints(width));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_WIDTH);
				}
				
				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_LEFT, PPTXUtils.convertEMUsToPoints(x));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_X);
				}

				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_TOP, PPTXUtils.convertEMUsToPoints(y));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_Y);
				}
				
				PPXMLImage ppxmlImage = (PPXMLImage) ppXMLObjectFactory.createObject(ppXML.PPXML_IMAGE);
				try 
				{
					ppxmlImage.setProperty(ppXML.PPXML_ATT_HEIGHT, PPTXUtils.convertEMUsToPoints(height));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_ATT_HEIGHT);
				}

				try 
				{
					ppxmlImage.setProperty(ppXML.PPXML_ATT_WIDTH, PPTXUtils.convertEMUsToPoints(width));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_ATT_WIDTH);
				}
				PPTXPicture pptxPicture = (PPTXPicture)this.pptxShape;
				if(Double.parseDouble(pptxPicture.getPptxBlipFill().getBottomOffset()) > 0.0)
				{
					try 
					{
						shape.setProperty(ppXML.PPXML_VALIGN_BOTTOM, pptxPicture.getPptxBlipFill().getBottomOffset());
					} catch (InvalidPropertyNameException e) 
					{
						DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_VALIGN_BOTTOM);
					}
				}
				if(Double.parseDouble(pptxPicture.getPptxBlipFill().getTopOffset()) > 0.0)
				{
					try 
					{
						shape.setProperty(ppXML.PPXML_VALIGN_TOP, pptxPicture.getPptxBlipFill().getTopOffset());
					} catch (InvalidPropertyNameException e) 
					{
						DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_VALIGN_TOP);
					}
				}

				if(Double.parseDouble(pptxPicture.getPptxBlipFill().getLeftOffset()) > 0.0)
				{
					try 
					{
						shape.setProperty(ppXML.PPXML_ALIGN_LEFT, pptxPicture.getPptxBlipFill().getLeftOffset());
					} catch (InvalidPropertyNameException e) 
					{
						DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_ALIGN_LEFT);
					}
				}
				if(Double.parseDouble(pptxPicture.getPptxBlipFill().getRightOffset()) > 0.0)
				{
					try 
					{
						shape.setProperty(ppXML.PPXML_ALIGN_RIGHT, pptxPicture.getPptxBlipFill().getRightOffset());
					} catch (InvalidPropertyNameException e) 
					{
						DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Image : " + ppXML.PPXML_ALIGN_RIGHT);
					}
				}

				String relationshipId = pptxPicture.getPptxBlipFill().getPptxBlip().getEmbed();

				InputStream imageStream = null;
				Iterator<String> iterator = pptxPicture.getImagesStreamsMap().keySet().iterator();
				while(iterator.hasNext())
				{
					String key = (String)iterator.next();
					if(key.contains(relationshipId))
					{
						imageStream = pptxPicture.getImagesStreamsMap().get(key);
						relationshipId = key;
						break;
					}
				}
				String fileName = context.getImageSaveDir().getAbsolutePath();
				if(imageStream != null)
				{					
					fileName = fileName + File.separator + relationshipId.substring(relationshipId.lastIndexOf("image"));
					File imageFile = new File (fileName);
					File imageSaveDir = context.getImageSaveDir();
					if(imageSaveDir.exists() && !imageFile.exists())
					{
						FileOutputStream fos = new FileOutputStream(fileName);
						int c;
				         while ((c = imageStream.read()) != -1) 
				         {
				        	 fos.write(c);
				         }
				         fos.flush();
				         fos.close();
				         imageStream.close();
				         this.pptxShape.getImagesStreamsMap().remove(relationshipId);
				         context.getImagesWrittinToDir().add(relationshipId);
					}
				}

				fileName = context.getImageSaveDir().getAbsolutePath();
				String href = fileName.substring(fileName.lastIndexOf(File.separator) + 1);
				href = href + File.separator + relationshipId.substring(relationshipId.lastIndexOf("image"));
				ppxmlImage.setProperty(ppXML.PPXML_ATT_HREF, href);

				this.shape.addChild(ppxmlImage);
				context.setImageCounter(context.getImageCounter() + 1);				
			}
			if(pptxShape != null && pptxShape instanceof PPTXShape && !(pptxShape instanceof PPTXPicture))
			{
				/*
				 * Shape Element's children
				 * 
				 * extLst (Extension List with Modification Flag)   Section 4.2.4
				 * p:nvSpPr (Non-Visual Properties for a Shape)     Section 4.4.1.31
				 * p:spPr (Shape Properties)                        Section 4.4.1.41 
				 * style (Shape Style)                              Section 4.4.1.43
				 * p:txBody (Shape Text Body)                       Section 4.4.1.47
				 */
				
				String noFill = String.valueOf(this.pptxShape.getShapeProperties().isNoFill());
				String outLineFillColor = this.pptxShape.getShapeProperties().getLnOutline().getSolidFill().getColorValue();
				try 
				{
					if(noFill != null && noFill.trim().length() > 0 && !(noFill.toLowerCase().equals("false")))
					{
						shape.setProperty(ppXML.PPXML_SHP_ATT_NO_FILL, noFill);
					}
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalid Property is set for ppxml Shape : " + ppXML.PPXML_SHP_ATT_NO_FILL);
				}

				try 
				{
					if(outLineFillColor != null && outLineFillColor.trim().length() > 0)
					{
						shape.setProperty(ppXML.PPXML_ATT_SHP_LINECOLOR, outLineFillColor);
					}
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalid Property is set for ppxml Shape : " + ppXML.PPXML_ATT_SHP_LINECOLOR);
				}
				
				
				int x = 0;
				int y = 0;
				int height = 0;
				int width = 0;
				String type = this.pptxShape.getPptxNonVisualProperties().getPptxPlaceholderShape().getType();
				
				if(type == null || type.equals(""))
				{
					type = "body";
				}

				if (this.pptxShape.getPptxConnectionNonVisualShapeProperties().isTxBox())
				{
					type = "other";
				}

				PPTXShape masterSlidePPTXShape = context.getMasterSlideShapeByType(type);
				
					x = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
					if(x < 0)
					{
						x = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
						if(x < 0)
						{
							if(masterSlidePPTXShape != null)
							{
								x = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getX());
							}
						}
					}
					y = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
					if(y < 0)
					{
						y = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
						if(y < 0)
						{
							if(masterSlidePPTXShape != null)
							{
								y = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getY());
							}
						}
					}
					
					width = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
					if(width < 0)
					{
						width = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
						if(width < 0)
						{
							if(masterSlidePPTXShape != null)
							{
								width = Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCx());
							}
						}
					}
					
					height = Integer.parseInt(this.pptxShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
					if(height < 0)
					{
						height = Integer.parseInt(this.pptxLayoutShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
						if(height < 0)
						{
							if(masterSlidePPTXShape != null)
							{
								height= Integer.parseInt(masterSlidePPTXShape.getShapeProperties().getPptx2DTransformforIndividualObjects().getCy());
							}
						}
					}
				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_HEIGHT, PPTXUtils.convertEMUsToPoints(height));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_HEIGHT);
				}

				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_WIDTH, PPTXUtils.convertEMUsToPoints(width));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_WIDTH);
				}
				
				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_LEFT, PPTXUtils.convertEMUsToPoints(x));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_X);
				}

				try 
				{
					shape.setProperty(ppXML.PPXML_ATT_TOP, PPTXUtils.convertEMUsToPoints(y));
				} catch (InvalidPropertyNameException e) 
				{
					DEBUG.errorPrint("PPTX - Slide: Invalide Property is set for ppxml Shape : " + ppXML.PPXML_ATT_Y);
				}

				PPXMLFrame ppxmlFrame = (PPXMLFrame) ppXMLObjectFactory.createObject(ppXML.PPXML_FRAME);
				ppxmlFrame.copyProperties(shape);
				
				boolean isTitleSyleApplied = false;
				
				List <PPTXTextParagraph> pptxTextParagraphList = pptxShape.getPptxTextParagraphList();
				for(int tpIndex = 0; tpIndex < pptxTextParagraphList.size(); tpIndex++)
				{
					PPTXTextParagraph pptxTextParagraph = pptxTextParagraphList.get(tpIndex);					
					 
					PPTXDefaultParagraphStyle pptxTextPara = pptxTextParagraph.getPara_TextParagraphProperties();

					PPXMLParagraph ppxmlParagraph = (PPXMLParagraph) ppXMLObjectFactory.createObject(ppXML.PPXML_PARAGRAPH);
					
					if( tpIndex == 0 )
					{
						if(context.getShapeNo() == 0 && context.getMarker_IDSlideNo() == 0)
						{
							PPXMLMarker ppXMLMarker_FirstSlide = (PPXMLMarker) ppXMLObjectFactory.createObject(ppXML.PPXML_MARKER);
							ppXMLMarker_FirstSlide.setName("_first");
							ppxmlParagraph.addChild(ppXMLMarker_FirstSlide);
						}
						
						if(context.getShapeNo()==0)
						{
							PPXMLMarker ppXMLMarker_SlideNo = (PPXMLMarker) ppXMLObjectFactory.createObject(ppXML.PPXML_MARKER);
							ppXMLMarker_SlideNo.setName("_slide"+context.getMarker_IDSlideNo());
							ppxmlParagraph.addChild(ppXMLMarker_SlideNo);
						}
						if(context.getShapeNo() == 0 && context.getMarker_IDSlideNo()== context.getMarker_LastSlideNo()-1)
						{
							PPXMLMarker ppXMLMarker_SlideNo = (PPXMLMarker) ppXMLObjectFactory.createObject(ppXML.PPXML_MARKER);
							ppXMLMarker_SlideNo.setName("_end");
							ppxmlParagraph.addChild(ppXMLMarker_SlideNo);
						}
					}
					
					//System.out.println("["+type+"]");
					
					PPXMLParaProperties ppxmlProperties = null;
					String masterSlide_FontSize = "";
					
					if(type.toLowerCase().contains("subtitle"))
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE,"center-body");
						
						if(pptxTextParagraph.getPara_TextParagraphProperties().getLevel() != null && !pptxTextParagraph.getPara_TextParagraphProperties().getLevel().equals(""))
						{
							ppxmlProperties = pptxTextParagraph.getPara_TextParagraphProperties().getBodyParaPropertiesByLevel();
						}
						else
						{
							ppxmlProperties = context.getPPTXSlideMasterTextStyles().getPptxOtherStyles().getpPTXListTextStyles().getLvl1pPr().getParagraphProperties();
						}
					}
					else if(type.toLowerCase().contains("title"))
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE,"title");
						ppxmlProperties = context.getPPTXSlideMasterTextStyles().getPptxTitleStyles().getpPTXListTextStyles().getLvl1pPr().getParagraphProperties();
					}
					else if(type.toLowerCase().contains("dt") || type.toLowerCase().contains("ftr") || type.toLowerCase().contains("sldnum"))
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE,"other");
						ppxmlProperties = masterSlidePPTXShape.getPptxListTextStyles().getLvl1pPr().getParagraphProperties();
					}
					else if(type.toLowerCase().equals("") || type.toLowerCase().contains("body"))
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE, "other");
						if(type.toLowerCase().contains("body"))
						{
							ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE, "body");
						}
						if(pptxTextParagraph.getPara_TextParagraphProperties().getLevel() != null && !pptxTextParagraph.getPara_TextParagraphProperties().getLevel().equals(""))
						{
							ppxmlProperties = pptxTextParagraph.getPara_TextParagraphProperties().getBodyParaPropertiesByLevel();
						}
						else
						{
							ppxmlProperties = context.getPPTXSlideMasterTextStyles().getPptxOtherStyles().getpPTXListTextStyles().getLvl1pPr().getParagraphProperties();
						}
					}
					else if (type.toLowerCase().contains("ctrTitle"))
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE,"center-title");
						ppxmlProperties = context.getPPTXSlideMasterTextStyles().getPptxTitleStyles().getpPTXListTextStyles().getLvl1pPr().getParagraphProperties();
					}
					else
					{
						ppxmlFrame.setProperty(ppXML.PPXML_ATT_TYPE,"other");
						ppxmlProperties = context.getPPTXSlideMasterTextStyles().getPptxOtherStyles().getpPTXListTextStyles().getLvl1pPr().getParagraphProperties();
					}

					masterSlide_FontSize = ppxmlProperties.getProperty(ppXML.PPXML_ATT_FONTSIZE);
					PPTXDefaultParagraphStyle pptxTextParaNew = pptxTextPara;
				
			       /**
					* Change started by <Shabana Majeed> on  <05/27/2011>
				    * Reason : <To override default allignment property with incoming paragraph align property>
					* fix for Issue # 25 --- The heading is centered aligned in xml but in source file it is left aligned. 
				    */
					
					String[][] paragraph_rPr_properties = 
						pptxTextParagraph.getPara_TextParagraphProperties().getParagraphProperties(true);
					
					for (int i = 0; i < paragraph_rPr_properties.length; i++)
					{
						if (paragraph_rPr_properties[i][0] != null && paragraph_rPr_properties[i][1] != null
								&& paragraph_rPr_properties[i][0].equals("align"))
						{
							ppxmlProperties.setProperty(paragraph_rPr_properties[i][0], paragraph_rPr_properties[i][1]);
															
						}
					}	
					
					/**
				     * Change ended by <Shabana Majeed> on  <05/27/2011>
				     */
					
					PPTXUtils.copyPropertiesToPPXMLObject(ppxmlProperties, ppxmlParagraph);												
					List<PPTXTextRun> pptxTextRunList = pptxTextParagraph.getPptxTextRunList();
					
					for(int trIndex = 0; trIndex < pptxTextRunList.size(); trIndex++)
					{
						PPTXTextRun pptxTextRun = pptxTextRunList.get(trIndex);
						
						// if there is a line break in input file then text run will come as br element
						if(pptxTextRun.isBreakRun())
						{
							PPXMLBreak ppxmlBreak = (PPXMLBreak) ppXMLObjectFactory.createObject(ppXML.PPXML_BREAK);
							ppxmlBreak.setProperty(ppXML.PPXML_ATT_TYPE,"line");
							ppxmlParagraph.addChild(ppxmlBreak);
							continue;
						}
						
						// if there is a text run with empty paragraph then it should not be displayed as 
						// its properties will be applied to next special text and properties of next run
						// will be overridden
						else if(pptxTextRun.getRun_Text().equals(""))
						{
							continue;
						}
						

						PPXMLText ppxmlSpecialText = (PPXMLText) ppXMLObjectFactory.createObject(ppXML.PPXML_SPECIALTEXT);
						PPXMLText ppxmlText = (PPXMLText) ppXMLObjectFactory.createObject(ppXML.PPXML_TEXT);
						
						PPXMLCharProperties ppxmlCharProperties =  pptxTextRun.getRun_TextRunProperties().getCharacterProperties();
						
						if(ppxmlCharProperties.getProperty(ppXML.PPXML_ATT_FONTSIZE).equals("0.0"))
						{
							ppxmlCharProperties.setProperty(ppXML.PPXML_ATT_FONTSIZE,masterSlide_FontSize);
						}

						if(trIndex==0)
						{
							pptxTextParaNew.setDefaultTextRunProperties(pptxTextRun.getRun_TextRunProperties());
							PPTXDefaultParagraphStyle rPr_pPr_Properties = new PPTXDefaultParagraphStyle();
							rPr_pPr_Properties.setDefaultTextRunProperties(pptxTextRun.getRun_TextRunProperties());
							
							String[][] rPr_properties = pptxTextParaNew.getParagraphProperties(true);
														
							for (int i= 0;i< rPr_properties.length;i++)
							{
								if (rPr_properties[i][0]!=null && rPr_properties[i][1]!=null)
								{
//									System.out.println("Coming Text run properties name ="+rPr_properties[i][0]);
//									System.out.println("Coming Text run properties value ="+rPr_properties[i][1]);
									ppxmlProperties.setProperty(rPr_properties[i][0],rPr_properties[i][1]);																	
								}
							}	
							
							String lineSpaceReduction = this.pptxShape.getPptxBodyProperties().getLnSpcReduction();
							double lineSpacing = PPTXUtils.convertFromStringToDouble(ppxmlProperties.getProperty(ppXML.PPXML_ATT_LINESPACING));
							
							if(lineSpaceReduction != null && !lineSpaceReduction.equals(""))
							{
								ppxmlProperties.setProperty(ppXML.PPXML_ATT_LINESPACING, 
										"" + (lineSpacing - (lineSpacing * (Double.parseDouble(lineSpaceReduction))) / 100.0));
							}
							
							if (!pptxTextPara.isNoBullet() && (
									(!(pptxTextPara.getBulletChar().equals("")) && !(pptxTextPara.getBulletChar()==null))
									|| !( pptxTextPara.getBulletFont().getPitchFamily().equals(""))
									|| (!( pptxTextPara.getBulletFont().getTypeface().equals("")) && 
											!( pptxTextPara.getBulletFont().getTypeface().equals("Cal")))
											))
							{
								if((pptxTextPara.getBulletChar() != null && !pptxTextPara.getBulletChar().equals(""))
									|| ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_CHARACTER) != null)
								{
									PPXMLListItemLabel listLabel = 
										(PPXMLListItemLabel) ppXMLObjectFactory.createObject(ppXML.PPXML_LIST_ITEM_LABEL);
									
									if(pptxTextPara.getBulletChar() != null && !pptxTextPara.getBulletChar().equals(""))
									{
										listLabel.setLabelText(pptxTextPara.getBulletChar());
									}									
									else
									{
										listLabel.setLabelText(ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_CHARACTER));
									}
									listLabel.setProperty(ppXML.PPXML_ATT_FONT,pptxTextPara.getBulletFont().getTypeface());
									
									if(Double.parseDouble(pptxTextPara.getBulletFontSize()) > 0)
									{
										listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE, pptxTextPara.getBulletFontSize());//""+(Double.parseDouble(ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONTSIZE))*(Double.parseDouble(pptxTextPara.getBuSzPct_Value())/1000.0))/100.0);
									}									
									else if(Double.parseDouble(pptxTextPara.getBulletFontSize()) == 0 && 
											Double.parseDouble(pptxTextPara.getBuSzPct_Value()) > 0)
									{
										listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE, 
												"" + (Double.parseDouble(
														 ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONTSIZE))
														 *
														 (Double.parseDouble(pptxTextPara.getBuSzPct_Value()) / 1000.0)) / 100.0);
										ppxmlParagraph.setProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE, 
												"" + (Double.parseDouble(
														ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONTSIZE))
														*
														(Double.parseDouble(pptxTextPara.getBuSzPct_Value()) / 1000.0)) / 100.0);							
									}
									
									else if(ppxmlProperties != null && 
											ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE) != null && 
											Double.parseDouble(ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE)) > 0)
									{
										if(Double.parseDouble(pptxTextPara.getBuSzPct_Value()) > 0)
										{
											ppxmlParagraph.setProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE, 
													"" + (Double.parseDouble(
															ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONTSIZE))
															*
															(Double.parseDouble(pptxTextPara.getBuSzPct_Value()) / 1000.0)) / 100.0);
										}
										listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE, 
												ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE));
									}
									else 
									{
										listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE,"22.4");
									}
									
									if(!pptxTextPara.getBuClr().getColorValue().equals(""))
									{
										listLabel.setProperty(ppXML.PPXML_ATT_FONTCOLOR, 
												"" + context.getColorValue(pptxTextPara.getBuClr().getColorValue()));
									}
									else 
									{
										listLabel.setProperty(ppXML.PPXML_ATT_FONTCOLOR, "#000000");//f07f09
									}
									ppxmlParagraph.addChild(listLabel);
								}
						}
							
//							
//							//!(pptxTextPara.getBulletChar()=="") || !( pptxTextPara.getBulletFont().getPitchFamily()=="") || ! (pptxTextPara.getBuSzPct_Value()=="") || ! (pptxTextPara.getBuClr().getColorValue()=="")
//							else if(!pptxTextPara.isNoBullet() && ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_CHARACTER) != null && 
//									ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE) != null &&
//									ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_COLOR) != null)
//							{
//								
//								System.out.println("comes in 2");
//								PPXMLListItemLabel listLabel = (PPXMLListItemLabel) ppXMLObjectFactory.createObject(ppXML.PPXML_LIST_ITEM_LABEL);
//								listLabel.setLabelText(ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_CHARACTER));
//								listLabel.setProperty(ppXML.PPXML_ATT_FONT,pptxTextPara.getBulletFont().getTypeface());
//								if(Double.parseDouble(ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE)) > 0.0)
//									listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE, ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_FONT_SIZE));
//								else
//									listLabel.setProperty(ppXML.PPXML_ATT_FONTSIZE, ppxmlProperties.getProperty(ppXML.PPXML_ATT_FONTSIZE));
//								listLabel.setProperty(ppXML.PPXML_ATT_FONTCOLOR, ppxmlProperties.getProperty(ppXML.PPXML_ATT_BULLET_COLOR));
//								ppxmlParagraph.addChild(listLabel);
//								
//							}
					
					  PPTXUtils.copyPropertiesToPPXMLObject(ppxmlProperties, ppxmlParagraph);
					}
						
						String fontScale = this.pptxShape.getPptxBodyProperties().getFontScale();
						double textRunFontSize = PPTXUtils.convertFromStringToDouble(ppxmlCharProperties.getProperty(ppXML.PPXML_ATT_FONTSIZE));

						if(fontScale != null && ! fontScale.equals(""))
						{
							ppxmlCharProperties.setProperty(ppXML.PPXML_ATT_FONTSIZE,
									(""+(textRunFontSize * (Double.parseDouble(fontScale))) / 100.0));
							ppxmlParagraph.setProperty(ppXML.PPXML_ATT_FONTSIZE,(""+(textRunFontSize * (Double.parseDouble(fontScale))) / 100.0));
							ppxmlProperties.setProperty(ppXML.PPXML_ATT_FONTSIZE, ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONTSIZE));
						}
						
					  if (pptxTextRun.getRun_TextRunProperties().getBaseLine()!="0")
						{
							double baseLine = Double.parseDouble(pptxTextRun.getRun_TextRunProperties().getBaseLine());
							if(baseLine<0)
							{
								ppxmlSpecialText.setProperty(ppXML.PPXML_ATT_EMPHASIS_SUBSCRIPT, "true");
							}
							else
							{
								ppxmlSpecialText.setProperty(ppXML.PPXML_ATT_EMPHASIS_SUPERSCRIPT, "true");
							}
							
							ppxmlSpecialText.setProperty(ppXML.PPXML_ATT_EMPHASIS_SCRIPTOFFSET, ""+baseLine/1000.0);
						}
				
					  	// if text run has some different properties than that of paragraph properties
					    // then it should be a special text node other wise it should be normal text node
					    // fix for Issue # 19 --- Special text tag for words having same attribute as paragraph
						if(ppxmlCharProperties!= null)
						{
							ppxmlCharProperties = PPTXUtils.getChangedProperties(ppxmlProperties, ppxmlCharProperties);
							if(ppxmlCharProperties.getPropertyCount() > 0)
							{
								PPTXUtils.copyPropertiesToPPXMLObject(ppxmlCharProperties, ppxmlSpecialText);
								ppxmlText = ppxmlSpecialText;
							}
							// if superscript is set then it should be a special text
							if(ppxmlSpecialText.getProperty(ppXML.PPXML_ATT_EMPHASIS_SUPERSCRIPT) .equals("true") ||
									(ppxmlSpecialText.getProperty(ppXML.PPXML_ATT_EMPHASIS_SUBSCRIPT) .equals("true")))
							{
								ppxmlText = ppxmlSpecialText;
							}
						}
											
						String text = pptxTextRun.getRun_Text();
						
						// if cap = all means all caps is applied at text run
						// fix for Issue # 16 ---- "All Caps" effect is not supported
						if(pptxTextRun.getRun_TextRunProperties().getCapitalization().equals("all"))
						{
							text = text.toUpperCase();
						}
						
						ppxmlText.setText(text);
						PPXMLLink hyperLink = (PPXMLLink) ppXMLObjectFactory.createObject(ppXML.PPXML_LINK);
						
					    /**
					     * Change started by <Shabana Majeed> on  <05/27/2011>
					     * Reason : <Add support for hyperlinks, jump to any slide, jump to url>
					     * fix for Issue # 7,8,10 and 11 (PPTX Driver | Local Build)
					     */
						
						if(! pptxTextRun.getRun_TextRunProperties().getHlinkClick().getHlinkClick_id().equals("") || 
								! pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().equals(""))
						{							
							if (! pptxTextRun.getRun_TextRunProperties().getHlinkClick().getHlinkClick_id().equals("")
									&& ! pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump"))
							{
								hyperLink.setProperty(ppXML.PPXML_ATT_HREF, 
										this.SlideRelationships.get(
												pptxTextRun.getRun_TextRunProperties().getHlinkClick().getHlinkClick_id()));
								hyperLink.addChild(ppxmlText);
								ppxmlParagraph.addChild(hyperLink);
							}
							
							if (! pptxTextRun.getRun_TextRunProperties().getHlinkMouseOver().getHlinkClick_id().equals(""))
							{
								hyperLink.setProperty(ppXML.PPXML_ATT_HREF,
										this.SlideRelationships.get(
												pptxTextRun.getRun_TextRunProperties().getHlinkMouseOver().getHlinkClick_id()));
								hyperLink.addChild(ppxmlText);
								ppxmlParagraph.addChild(hyperLink);
							}
							
							if (! pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().equals(""))
							{
	
								if(pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=nextslide"))
								{
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, "#_slide" + (context.getMarker_IDSlideNo() + 1));
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=previousslide"))
								{
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, "#_slide" + (context.getMarker_IDSlideNo() - 1));
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("hlinksldjump")
										&& 
										!(pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=firstslide"))
										)
								{
									String jumpToSlide = this.SlideRelationships.get(
											pptxTextRun.getRun_TextRunProperties().getHlinkClick().getHlinkClick_id());
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, 
											"#_slide"
											+
											(Integer.parseInt(jumpToSlide.replace(".xml", "").substring(5)) - 1)
											);
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=firstslide"))
								{
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, "#_slide0");
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=lastslide"))
								{
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, "#_end");
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=lastslideviewed"))
								{
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
								
								else if (pptxTextRun.getRun_TextRunProperties().getHlinkClick().getAction().contains("jump=endshow"))
								{
									hyperLink.setProperty(ppXML.PPXML_ATT_HREF, "#_end");
									hyperLink.addChild(ppxmlText);
									ppxmlParagraph.addChild(hyperLink);
								}
							}
						}
						
					    /**
					     * Change ended by <Shabana Majeed> on  <05/27/2011>
					   */
						
						else
						{
							ppxmlParagraph.addChild(ppxmlText);
						}
						
						PPXMLMarker ppXMLMarker_SlideNo = (PPXMLMarker) ppXMLObjectFactory.createObject(ppXML.PPXML_MARKER);
						ppXMLMarker_SlideNo.setName("_slide"+context.getMarker_IDSlideNo());
						
					}// end of for loop for text runs
					
				// apply end paragraph properties to paragraph when there is no text run

				if( pptxTextRunList.size() == 0 && type.equals("body"))
				{
					PPXMLCharProperties endPara_CharacterProp = pptxTextParagraph.getEndParaRPr_textRunProperties().getCharacterProperties();
					
					String fontScale = this.pptxShape.getPptxBodyProperties().getFontScale();
					double textRunFontSize = PPTXUtils.convertFromStringToDouble(ppxmlProperties.getProperty(ppXML.PPXML_ATT_FONTSIZE));
					
					if(fontScale != null && ! fontScale.equals(""))
					{
						endPara_CharacterProp.setProperty(ppXML.PPXML_ATT_FONTSIZE,(""+(textRunFontSize * (Double.parseDouble(fontScale)))/100.0));
					}
					
					PPTXUtils.copyPropertiesToPPXMLObject(endPara_CharacterProp, ppxmlParagraph);					
				}
					
                   if(pptxTextParagraph.isFieldPresent())
					{
						PPXMLText ppxmlText = (PPXMLText) ppXMLObjectFactory.createObject(ppXML.PPXML_SPECIALTEXT);
						PPXMLCharProperties ppxmlCharProperties =  pptxTextParagraph.getPptxTextField().getFld_TextRunProperties().getCharacterProperties();
						if(ppxmlCharProperties != null)
						{
							ppxmlCharProperties = PPTXUtils.getChangedProperties(ppxmlProperties, ppxmlCharProperties);
							if(ppxmlCharProperties.getPropertyCount() > 0)
							{					
								PPTXUtils.copyPropertiesToPPXMLObject(ppxmlCharProperties, ppxmlText);
							}
						}
						String text = pptxTextParagraph.getPptxTextField().getFld_Text();
						ppxmlText.setText(text);
						ppxmlParagraph.addChild(ppxmlText);

					}
                   
                   /**
                    * Change started by <Shabana Majeed> on  <06/07/2011>
                    * Reason : <Font "calibri" is just "cal" in pptx>
                    * fix for Issue # 29 (PPTX Driver | Local Build)
                    */

                   if(ppxmlParagraph.getProperty(ppXML.PPXML_ATT_FONT).equals("Cal"))
                   {
                	   ppxmlParagraph.setProperty(ppXML.PPXML_ATT_FONT, "Calibri");
                   }
                   
                   if(ppxmlParagraph.getProperty(ppXML.PPXML_ATT_BULLET_FONT).equals("Cal"))
                   {
                	   ppxmlParagraph.setProperty(ppXML.PPXML_ATT_BULLET_FONT, "Calibri");
                   }

                   /**
                    * Change ended by <Shabana Majeed> on  <06/07/2011>
                    */
                   
                   ppxmlFrame.addChild(ppxmlParagraph);
				}
				
				shape.addChild(ppxmlFrame);
				
//				NodeList spChildren = shapeElement.getChildNodes();
//
//				for(int index = 0; index < spChildren.getLength(); index++)
//				{
//					Node spChild = spChildren.item(index);
//					Element spChildElement = (Element) spChild;
//					String spChildElementName = spChildElement.getNodeName();
//					
//					if(spChildElementName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDE_SHAPE_NON_VISUAL_PROPERTIES))
//					{
//						Element nonVisualProperties = (Element) spChildElement.
//						getElementsByTagName(PPTXConstants.PPTX_SLIDE_SHAPE_NON_VISUAL_DRAWING_PROPERTIES).item(0);
//						if(nonVisualProperties != null)
//						{
//							String id = nonVisualProperties.getAttribute(PPTXConstants.PPTX_ATT_SLIDE_SHAPE_NON_VISUAL_PROPERTY_ID);
//							if(id != null && id.trim().length() > 0)
//							{
//								shapeId = Integer.parseInt(id);
//							}// end if (id != null)
//						}					
//
//					}
//					else if(spChildElementName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDE_SHAPE_PROPERTIES))
//					{						
//						parseShapeProperties(spChildElement, shapeId);
//					}
//					else if(spChildElementName.equalsIgnoreCase(PPTXConstants.PPTX_SLIDE_SHAPE_TEXT_BODY))
//					{
//						Element TXBodyOfLayout = null;
//						if(shapeElementOfLayout != null)
//						{
//							TXBodyOfLayout = (Element)shapeElementOfLayout.getElementsByTagName(PPTXConstants.PPTX_SLIDE_SHAPE_TEXT_BODY).item(0);
//						}
//						PPXMLFrame ppxmlFrame = null;
//						if(TXBodyOfLayout != null)
//						{
//							ppxmlFrame = processPPTX_TXBody(spChildElement, TXBodyOfLayout);
//						}
//						else
//						{
//							ppxmlFrame = processPPTX_TXBody(spChildElement, null);
//						}
//						
//						ppxmlFrame.copyProperties(shape);
//						shape.addChild(ppxmlFrame);
//					}						
//				}
			}
		}
		catch(Exception e)
		{
			BlockerException blockerExp = new BlockerException("PPTX - Error while parsing Shape Element : - " + e.getMessage());
			blockerExp.setStackTrace(e.getStackTrace());
			throw blockerExp;
		}
		return true;
	}

	/**
	 * @see pptx.parser.Parser#getOutputData()
	 */
	public ParserOutputData getOutputData() throws BlockerException 
	{
		if(shape != null)
		{
			ParserOutputData parserOutputData = new ParserOutputData();
			parserOutputData.setShape(shape);
			return parserOutputData;
		}
		return null;
	}
	
	private PPXMLFrame processPPTX_TXBody(Element txBodyElement, Element txBodyElementOfLayout) throws BlockerException
	{
		PPXMLFrame ppxmlFrame = null;
		Parser parser = parserFactory.createParserObject(PPTXConstants.PPTX_TEXTBODYPARSER_CLASS_NAME);
    	
    	ParserInputData parserInputData = new ParserInputData();
    	parserInputData.setTXBodyElement(txBodyElement);
    	parserInputData.setTXBodyElementOfLayout(txBodyElementOfLayout);
    	parser.populateData(parserInputData);

    	boolean isParsingComplete = parser.parse();
    	
    	if(isParsingComplete)
    	{
    		ppxmlFrame = parser.getOutputData().getPPPXMLFrame();
    	}

		return ppxmlFrame;
	}
	

}
