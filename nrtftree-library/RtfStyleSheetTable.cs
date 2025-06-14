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
 * Class:		RtfStyleSheetTable
 * Description:	Represents the style sheet table of an RTF document.
 * ******************************************************************************/

using System.Collections;

namespace Net.Sgoliver.NRtfTree.Util
{
    /// <summary>
    /// Represents the style sheet table of an RTF document.
    /// </summary>
    public class RtfStyleSheetTable
    {
        private Dictionary<int, RtfStyleSheet> stylesheets = null;

        /// <summary>
        /// Constructor.
        /// </summary>
        public RtfStyleSheetTable()
        {
            stylesheets = new Dictionary<int, RtfStyleSheet>();
        }

        /// <summary>
        /// Add a new style to the style table. The style will be added with a new index not existing in the table.
        /// </summary>
        /// <param name="ss">New style to add to the table.</param>
        public void AddStyleSheet(RtfStyleSheet ss)
        {
            ss.Index = newStyleSheetIndex();

            stylesheets.Add(ss.Index, ss);
        }

        /// <summary>
        /// Add a new style to the style table. The style will be added at the index.
        /// </summary>
        /// <param name="index">Index at which to add the stylesheet.</param>
        /// <param name="ss">New stylesheet to add to the table.</param>
        public void AddStyleSheet(int index, RtfStyleSheet ss)
        {
            ss.Index = index;

            stylesheets.Add(index, ss);
        }

        /// <summary>
        /// Remove a stylesheet at the index.
        /// </summary>
        /// <param name="index">Index at which to remove stylesheet</param>
        public void RemoveStyleSheet(int index)
        {
            stylesheets.Remove(index);
        }

        /// <summary>
        /// Remove the referenced stylesheet from table.
        /// </summary>
        /// <param name="ss">Stylesheet reference to remove</param>
        public void RemoveStyleSheet(RtfStyleSheet ss)
        {
            stylesheets.Remove(ss.Index);
        }

        /// <summary>
        /// Get a stylesheet by index.
        /// </summary>
        /// <param name="index">Index to get stylesheet from.</param>
        /// <returns>Returns stylesheet reference from the index in the table.</returns>
        public RtfStyleSheet GetStyleSheet(int index)
        {
            return stylesheets[index];
        }

        /// <summary>
        /// Get a stylesheet by index
        /// </summary>
        /// <param name="index">Index at which to get stylesheet.</param>
        /// <returns>Reference to stylesheet at index.</returns>
        public RtfStyleSheet this[int index]
        {
            get
            {
                return stylesheets[index];
            }
        }

        /// <summary>
        /// Number of stylesheets in table
        /// </summary>
        public int Count
        {
            get
            {
                return stylesheets.Count;
            }
        }

        /// <summary>
        /// Index of stylesheet referenced by name.
        /// </summary>
        /// <param name="name">Name of stylesheet.</param>
        /// <returns>Index of stylesheet with given name.</returns>
        public int IndexOf(string name)
        {
            int intIndex = -1;
            IEnumerator fntIndex = stylesheets.GetEnumerator();

            fntIndex.Reset();
            while (fntIndex.MoveNext())
            {
                if (((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Value.Name.Equals(name))
                {
                    intIndex = (int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key;
                    break;
                }
            }

            return intIndex;
        }

        /// <summary>
        /// Calculate a new free index in the table to add to.
        /// </summary>
        /// <returns>Index free to add at.</returns>
        private int newStyleSheetIndex()
        {
            int intIndex = -1;
            IEnumerator fntIndex = stylesheets.GetEnumerator();

            fntIndex.Reset();
            while (fntIndex.MoveNext())
            {
                if ((int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key > intIndex)
                    intIndex = (int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key;
            }

            return (intIndex + 1);
        }
    }
    
}
