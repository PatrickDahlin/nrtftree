/********************************************************************************
 *   This file is part of NRtfTree Library.
 *
 *   NRtfTree Library is free software; you can redistribute it and/or modify
 *   it under the terms of the GNU Lesser General Public License as published by
 *   the Free Software Foundation; either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   NRtfTree Library is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU Lesser General Public License for more details.
 *
 *   You should have received a copy of the GNU Lesser General Public License
 *   along with this program. If not, see <http://www.gnu.org/licenses/>.
 ********************************************************************************/

/********************************************************************************
 * Library:		NRtfTree
 * Version:     v0.4
 * Date:		29/06/2013
 * Copyright:   2006-2013 Salvador Gomez
 * Home Page:	http://www.sgoliver.net
 * GitHub:	    https://github.com/sgolivernet/nrtftree
 * Class:		RtfPullParser
 * Description:	Pull parser of RTF documents.
 * ******************************************************************************/


namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Pull parser of RTF documents.
/// </summary>
public class RtfPullParser
{
    #region Constants

    public const int START_DOCUMENT = 0;
    public const int END_DOCUMENT = 1;
    public const int KEYWORD = 2;
    public const int CONTROL = 3;
    public const int START_GROUP = 4;
    public const int END_GROUP = 5;
    public const int TEXT = 6;

    #endregion

    #region Fields

    private TextReader? rtf;		//File/text input RTF
    private RtfLex? lex;		    //Lexical analyzer of RTF
    private RtfToken? tok;		//Token
    private int currentEvent;   //Event

    #endregion

    #region Constructors

    /// <summary>
    /// Constructor.
    /// </summary>
    public RtfPullParser()
    {
        currentEvent = START_DOCUMENT;
    }

    /// <summary>
    /// Load a file in RTF format
    /// </summary>
    /// <param name="path">Path to the file.</param>
    public int LoadRtfFile(string path)
    {
        // Initialize input
        rtf = new StreamReader(path);

        // Initialize lexical analyzer
        lex = new RtfLex(rtf);

        return 0;
    }

    /// <summary>
    /// Load an RTF document from text stream
    /// </summary>
    /// <param name="text">Text containing the rtf document.</param>
    public int LoadRtfText(string text)
    {
        // Initialize input
        rtf = new StringReader(text);

        // Initialize lexical analyzer
        lex = new RtfLex(rtf);

        return 0;
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Get the current event type.
    /// </summary>
    /// <returns>Type of event</returns>
    public int GetEventType()
    {
        return currentEvent;
    }

    /// <summary>
    /// Gets the next element in the document.
    /// </summary>
    /// <returns>Next element of the document.</returns>
    public int Next()
    {
        tok = lex?.NextToken();

        switch (tok?.Type)
        {
            case RtfTokenType.GroupStart:
                currentEvent = START_GROUP;
                break;
            case RtfTokenType.GroupEnd:
                currentEvent = END_GROUP;
                break;
            case RtfTokenType.Keyword:
                currentEvent = KEYWORD;
                break;
            case RtfTokenType.Control:
                currentEvent = CONTROL;
                break;
            case RtfTokenType.Text:
                currentEvent = TEXT;
                break;
            case RtfTokenType.Eof:
                currentEvent = END_DOCUMENT;
                break;
        }

        return currentEvent;
    }

    /// <summary>
    /// Gets the keyword/control symbol of the current element.
    /// </summary>
    /// <returns>Keyword/control symbol of the current element.</returns>
    public string? GetName()
    {
        return tok?.Key;
    }

    /// <summary>
    /// Get the parameter of the current element
    /// </summary>
    /// <returns>Parameter value of current element.</returns>
    public int? GetParam()
    {
        return tok?.Parameter;
    }

    /// <summary>
    /// Check if current element has a parameter value.
    /// </summary>
    /// <returns>Returns true if current element has a parameter.</returns>
    public bool? HasParam()
    {
        return tok?.HasParameter;
    }

    /// <summary>
    /// Get the text of the current element.
    /// </summary>
    /// <returns>Text of current element.</returns>
    public string? GetText()
    {
        return tok?.Key;
    }

    #endregion
}

