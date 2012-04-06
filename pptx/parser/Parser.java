
package pptx.parser;

import pptx.exception.BlockerException;
import pptx.parser.ParserInputData;
/**
 * @author sheraz.ahmed
 *
 */
public interface Parser {

	public abstract void populateData(ParserInputData parserInputData) throws BlockerException;
	public abstract boolean parse() throws BlockerException;
	public abstract ParserOutputData getOutputData() throws BlockerException;
}
