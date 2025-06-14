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
 * Class:		RtfFontTable
 * Description:	Table of Fonts of an RTF document.
 * ******************************************************************************/

using System.Collections;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Table of Fonts of an RTF document.
        /// </summary>
        public class RtfFontTable
        {
            /// <summary>
            /// Internal list of fonts.
            /// </summary>
            private Dictionary<int,string> fonts;

            /// <summary>
            /// Constructor
            /// </summary>
            public RtfFontTable()
            {
                fonts = new Dictionary<int,string>();
            }

            /// <summary>
            /// Insert a new font to the table.
            /// </summary>
            /// <param name="name">New font to insert.</param>
            public void AddFont(string name)
            {
                fonts.Add(newFontIndex(),name);
            }

            /// <summary>
            /// Insert a new font to the table.
            /// </summary>
            /// <param name="index">Index to insert font to.</param>
            /// <param name="name">New font to insert.</param>
            public void AddFont(int index, string name)
            {
                fonts.Add(index, name);
            }

            /// <summary>
            /// Obtain the font at index.
            /// </summary>
            /// <param name="index">Index of the font to retrieve.</param>
            /// <returns>Font at the n-th position in the table.</returns>
            public string this[int index]
            {
                get
                {
                    return fonts[index];
                }
            }

            /// <summary>
            /// Number of fonts in the table.
            /// </summary>
            public int Count
            {
                get 
                {
                    return fonts.Count;
                }
            }

            /// <summary>
            /// Obtain the index of a specific font.
            /// </summary>
            /// <param name="name">Font to find.</param>
            /// <returns>Index of font in the table.</returns>
            public int IndexOf(string name)
            {
                int intIndex = -1;
                IEnumerator fntIndex = fonts.GetEnumerator();

                fntIndex.Reset();
                while (fntIndex.MoveNext())
                {
                    if (((KeyValuePair<int,string>)fntIndex.Current).Value.Equals(name))
                    {
                        intIndex = (int)((KeyValuePair<int, string>)fntIndex.Current).Key;
                        break;
                    }
                }

                return intIndex;
            }

            /// <summary>
            /// Get next free font index in the table.
            /// </summary>
            /// <returns>Next free index.</returns>
            private int newFontIndex()
            {
                int intIndex = -1;
                IEnumerator fntIndex = fonts.GetEnumerator();

                fntIndex.Reset();
                while (fntIndex.MoveNext())
                {
                    if ((int)((KeyValuePair<int, string>)fntIndex.Current).Key > intIndex)
                        intIndex = (int)((KeyValuePair<int, string>)fntIndex.Current).Key;
                }

                return (intIndex + 1);
            }
        }
    }
}
