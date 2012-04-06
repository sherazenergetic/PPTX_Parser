package pptx.exception;

public class BlockerException extends Exception 
{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public BlockerException() 
	{
		super();
	}

	public BlockerException(String message) 
	{
		super(message);
	}

	public BlockerException(String message, Throwable cause) 
	{
		super(message, cause);
	}

	public BlockerException(Throwable cause) 
	{
		super(cause);
	}

}
