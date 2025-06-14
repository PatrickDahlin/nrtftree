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
 * Class:		RtfLex
 * Description:	Document Lexical Analyzer of RTF.
 * ******************************************************************************/

using System.Text;

namespace Net.Sgoliver.NRtfTree.Core
{

    /// <summary>
    /// Lexical analyzer (tokenizer) for documents in RTF format. Analyze the document and return
    /// all RTF elements read (tokens).
    /// </summary>
    public class RtfLex
    {
        #region Private fields

        /// <summary>
        /// The open file.
        /// </summary>
        private TextReader rtf;

        /// <summary>
        /// Stringbuilder to build the keyword / text fragment
        /// </summary>
        private StringBuilder keysb;

        /// <summary>
        /// Stringbuilder to build the keyword parameter
        /// </summary>
        private StringBuilder parsb;

        /// <summary>
        /// Character read of the document
        /// </summary>
        private int c;

        #endregion

        #region Constantes

        /// <summary>
        /// Marker for end of file
        /// </summary>
        private const int Eof = -1;

        #endregion

        #region Constructores

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rtfReader">Filestream to parse</param>
        public RtfLex(TextReader rtfReader)
        {
            rtf = rtfReader;

            keysb = new StringBuilder();
            parsb = new StringBuilder();

            // The first character of the document is read
            c = rtf.Read();
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Read a new token from the RTF document.
        /// </summary>
        /// <returns>Siguiente token leido del documento.</returns>
        public RtfToken NextToken()
        {
            //Se crea el nuevo token a devolver
            RtfToken token = new RtfToken();

            // Carriage returns, tabulators and null characters are ignored
            while (c == '\r' || c == '\n' || c == '\t' || c == '\0')
                c = rtf.Read();

            // The character read is interpreted
            if (c != Eof)
            {
                switch (c)
                {
                    case '{':
                        token.Type = RtfTokenType.GroupStart;
                        c = rtf.Read();
                        break;
                    case '}':
                        token.Type = RtfTokenType.GroupEnd;
                        c = rtf.Read();
                        break;
                    case '\\':
                        parseKeyword(token);
                        break;
                    default:
                        token.Type = RtfTokenType.Text;
                        parseText(token);
                        break;
                }
            }
            else
            {
                token.Type = RtfTokenType.Eof;
            }

            return token;
        }


        #endregion

        #region Private Methods

        /// <summary>
        /// Read a key word of the RTF document.
        /// </summary>
        /// <param name="token">RTF Token to which the keyword will be assigned.</param>
        private void parseKeyword(RtfToken token)
        {
            // StringBuilders are reset
            keysb.Length = 0;
            parsb.Length = 0;

            int tokenParameter = 0;
            bool negative = false;

            c = rtf.Read();

            //If the character read is not a letter -> It is a control symbol or a special character: '\\', '\{' o '\}'
            if (!Char.IsLetter((char)c))
            {
                if (c == '\\' || c == '{' || c == '}')  // Special character
                {
                    token.Type = RtfTokenType.Text;
                    token.Key = ((char)c).ToString();
                }
                else // Control symbol
                {
                    token.Type = RtfTokenType.Control;
                    token.Key = ((char)c).ToString();

                    //If it is a special character (8 -bit code) the hexadecimal parameter is read
                    if (token.Key == "\'")
                    {
                        string cod = "";

                        cod += (char)rtf.Read();
                        cod += (char)rtf.Read();

                        token.HasParameter = true;

                        token.Parameter = Convert.ToInt32(cod, 16);
                    }

                    //TODO: Are there more control symbols with parameters?
                }

                c = rtf.Read();
            }
            else  // The read character is a letter
            {
                // The complete keyword is read (until you find a non -alphanumeric character, for example '\' or '' '
                while (Char.IsLetter((char)c))
                {
                    keysb.Append((char)c);

                    c = rtf.Read();
                }

                // The keyword read is assigned
                token.Type = RtfTokenType.Keyword;
                token.Key = keysb.ToString();

                // Check if the keyword has parameter
                if (Char.IsDigit((char)c) || c == '-')
                {
                    token.HasParameter = true;

                    // Check if the parameter is negative
                    if (c == '-')
                    {
                        negative = true;

                        c = rtf.Read();
                    }

                    // Read the parameter to the end
                    while (Char.IsDigit((char)c))
                    {
                        parsb.Append((char)c);

                        c = rtf.Read();
                    }

                    tokenParameter = Convert.ToInt32(parsb.ToString());

                    if (negative)
                        tokenParameter = -tokenParameter;

                    //Se asigna el par�metro de la palabra clave
                    token.Parameter = tokenParameter;
                }

                if (c == ' ')
                {
                    c = rtf.Read();
                }
            }
        }

        /// <summary>
        /// Read a text stream of the RTF document.
        /// </summary>
        /// <param name="token">RTF Token to which the keyword will be assigned.</param>
        private void parseText(RtfToken token)
        {
            // Reset StringBuilder
            keysb.Length = 0;

            while (c != '\\' && c != '}' && c != '{' && c != Eof)
            {
                keysb.Append((char)c);

                c = rtf.Read();

                // Carriage returns, tabs and null characters are ignored
                while (c == '\r' || c == '\n' || c == '\t' || c == '\0')
                    c = rtf.Read();
            }

            token.Key = keysb.ToString();
        }

        #endregion
    }
    
}
