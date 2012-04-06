/**
 * 
 */
package pptx.parser;

import java.util.Vector;




/**
 * @author sheraz.ahmed
 *
 */
public class PPTXParserFactory {
	
	private static Vector<PPTXParserFactory> mOutputMappings = new Vector<PPTXParserFactory>();	
	private Thread conversionThread = null;
	
	private PPTXParserFactory()
	{
		
	}
	/**
	 * 
	 * @param ctx
	 * @return
	 */
	private static boolean addOutputMapping(PPTXParserFactory ctx)
	{
		mOutputMappings.add(ctx);
		return true;
	}
	
	/**
	 * 
	 * @param conversionThread
	 */
	private static void removeOutputMapping(Thread param_conversionThread)
	{	
		PPTXParserFactory factory = null;

		for (int i = 0; i < mOutputMappings.size(); i++)
		{
			factory = mOutputMappings.get(i);
			if (factory.getConversionThread().equals(param_conversionThread))
			{
				mOutputMappings.remove(i);
				return;
			}				
		}		
	}
	
	/**
	 * 
	 * @param conversionThread
	 * @return
	 */
	private static PPTXParserFactory getOutputMapping(Thread param_conversionThread)
	{	
		PPTXParserFactory factory = null;
		PPTXParserFactory result = null;

		for (int i = 0; i < mOutputMappings.size(); i++)
		{
			factory = mOutputMappings.get(i);
			if (factory.getConversionThread().equals(param_conversionThread))
			{
				result = factory;
				break;
			}				
		}

		return result;
	}	
	
	/**
	 * 
	 * @return
	 */
	public static PPTXParserFactory getInstance()
	{
		Thread currThread = Thread.currentThread();
		PPTXParserFactory factory = getOutputMapping(currThread);


		if(factory == null)
		{
			factory = new PPTXParserFactory();
			factory.setConversionThread(currThread);
			addOutputMapping(factory);
		}
		return factory;
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
	 * @return
	 */
	public Thread getConversionThread() {
		return conversionThread;
	}
	
	/**
	 * 
	 * @param conversionThread
	 */
	public void setConversionThread(Thread param_conversionThread) {
		this.conversionThread = param_conversionThread;
	}
	
	public Parser createParserObject(String parserClassName)
	{
		try
		{
			ClassLoader classLoader = ClassLoader.getSystemClassLoader();
			Class<?> clazz = classLoader.loadClass(parserClassName);
			Object classObject = clazz.newInstance();
			return (Parser)classObject;
		}catch(ClassNotFoundException e)
		{
			DEBUG.errorPrint("PPTX - Class is not loadded " + parserClassName 
					+ ", and exception message=" + e.getMessage());
		} catch (InstantiationException e) {
			DEBUG.errorPrint("PPTX - Class " + parserClassName 
					+ " is not instantiated, and exception message=" + e.getMessage());
		} catch (IllegalAccessException e) {
			DEBUG.errorPrint("PPTX - Class access for " + parserClassName 
					+ " is illegal here, and exception message=" + e.getMessage());
		}
		return null;
	}
	
	

}
