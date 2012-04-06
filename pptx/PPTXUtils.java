/**
 * 
 */
package pptx;

import java.text.DecimalFormat;

import pptx.exception.BlockerException;



/**
 * @author sheraz.ahmed
 *
 */
public class PPTXUtils 
{
	public static double convertPercentToPoints(double percentValue)
	{
		double convertionConstant = 0.12 ;     //percentage convertion constant in points "http://www.metahead.com/converting-font-sizes-from-pixel-point-em-percent"
		return (percentValue * convertionConstant);
	}

	public static double convertFromStringToDouble (String stringValue) throws BlockerException
	{
		double convertedDouble = 0.0;
		try
		{
			convertedDouble = Double.parseDouble(stringValue);
		}
		catch (Exception e) {
			
			BlockerException blockerExp = new BlockerException("Error while converting string to double" + e.getMessage());
    		blockerExp.setStackTrace(e.getStackTrace());
    		throw blockerExp;
		}
		return convertedDouble;
	}
	
	public static String convertEMUsToPoints(int EMUValue)
	{
		DecimalFormat df = new DecimalFormat("0.00");
		double leftMarginDoubleValue = ((long)EMUValue * (double)72)/(double)914400;
		return df.format(leftMarginDoubleValue);
	}
	
	public static void copyPropertiesToPPXMLObject(PPXMLProperties ppxmlProperties, PPXMLObject ppxmlObject) throws Exception 
	{
		if(ppxmlProperties != null)
		{
			String validNames[] = ppxmlProperties.getValidNames();
			for(int i = 0; i < validNames.length; i++)
			{
				String propertyName = validNames[i];
				String propertyValue = ppxmlProperties.getProperty(propertyName);
				if(propertyValue != null && propertyValue.trim().length() > 0)
				{
					ppxmlObject.setProperty(propertyName, propertyValue);
				}
			}
		}
	}
	
	public static PPXMLCharProperties getChangedProperties(
			PPXMLParaProperties ppxmlParaProperties, 
			PPXMLCharProperties ppxmlCharPropertiesOriginal) throws Exception
	{
		PPXMLCharProperties ppxmlCharProperties = new PPXMLCharProperties();
		if(ppxmlParaProperties != null && ppxmlCharPropertiesOriginal != null)
		{
			String validNames[] = ppxmlCharPropertiesOriginal.getValidNames();
			for(int i = 0; i < validNames.length; i++)
			{
				String propertyName = validNames[i];
				String propertyValue = ppxmlCharPropertiesOriginal.getProperty(propertyName);
				String propertyValue2 = ppxmlParaProperties.getProperty(propertyName);
				if(propertyValue != null                  && propertyValue2 != null             &&
						propertyValue.trim().length() > 0 && propertyValue2.trim().length() > 0 &&
						!propertyValue.equals("0")        && !propertyValue2.equals("0")        &&
						!propertyValue.equals("0.0")      && !propertyValue2.equals("0.0")      &&
                        !propertyValue.equals("Cal")      && !propertyValue2.equals("Cal")      &&
						!propertyValue.equalsIgnoreCase(propertyValue2) )
				{
					ppxmlCharProperties.setProperty(propertyName, propertyValue);
				}			
			}
		}
		return ppxmlCharProperties;
	}
}
