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
 * Class:		RtfReader
 * Description:	Sequentially analyze RTF Dcoument.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// This class provides the methods necessary for sequential loading and parsing of an RTF document.
/// </summary>
public class RtfReader
{
    #region Private Fields

    private TextReader? rtf;		//File/text input reader RTF
    private RtfLex? lex;		//Lexical analyzer of RTF
    private RtfToken? tok;		//Token
    private SarParser? reader;		//Rtf Reader

    #endregion


    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="reader">
    /// Object of the SARParser type that contains the methods necessary for processing the different elements of an RTF document.
    /// </param>
    public RtfReader(SarParser reader)
    {
        this.reader = reader;
    }


    #region Public Methods

    /// <summary>
    /// Loads an RTF document given the path of the file that contains it.
    /// </summary>
    /// <param name="path">Path of the file containing the RTF document.</param>
    /// <returns>
    /// Result of the operation, 0 if successful
    /// </returns>
    public int LoadRtfFile(string path)
    {
        // Input is prepared
        rtf = new StreamReader(path);

        // Initialize lexical analyzer
        lex = new RtfLex(rtf);

        return 0;
    }

    /// <summary>
    /// Loads an RTF document given the character string that contains it.
    /// </summary>
    /// <param name="text">Character string containing the RTF document.</param>
    /// <returns>
    /// Result of the loading, 0 if successful
    /// </returns>
    public int LoadRtfText(string text)
    {
        // Input is prepared
        rtf = new StringReader(text);

        // Initialize lexical analyzer
        lex = new RtfLex(rtf);

        return 0;
    }

    /// <summary>
    /// Begin the analysis of the RTF document and calls the different methods of the IRtfReader object indicated in the class constructor.
    /// </summary>
    /// <returns>
    /// Result of the parsing, 0 is successful
    /// </returns>
    public int Parse()
    {
        // Resultcode
        var res = 0;

        // Begin a new document
        reader?.StartRtfDocument();

        // Get the first token
        tok = lex?.NextToken();

        while (tok != null && tok.Type != RtfTokenType.Eof)
        {
            switch (tok.Type)
            {
                case RtfTokenType.GroupStart:
                    reader?.StartRtfGroup();
                    break;
                case RtfTokenType.GroupEnd:
                    reader?.EndRtfGroup();
                    break;
                case RtfTokenType.Keyword:
                    reader?.RtfKeyword(tok.Key, tok.HasParameter, tok.Parameter);
                    break;
                case RtfTokenType.Control:
                    reader?.RtfControl(tok.Key, tok.HasParameter, tok.Parameter);
                    break;
                case RtfTokenType.Text:
                    reader?.RtfText(tok.Key);
                    break;
                default:
                    res = -1;
                    break;
            }

            // Get next singular token
            tok = lex?.NextToken();
        }

        //Finalize doucment
        reader?.EndRtfDocument();

        rtf?.Close();

        return res;
    }

    #endregion
}
