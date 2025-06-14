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
 * Class:		RtfColorTable
 * Description:	Table of colors in a RTF document.
 * ******************************************************************************/

using System.Drawing;

namespace Net.Sgoliver.NRtfTree.Util;

/// <summary>
/// Table of colors in a RTF document.
/// </summary>
public class RtfColorTable
{
    /// <summary>
    /// Internal list of colors
    /// </summary>
    private List<int> colors;

    /// <summary>
    /// Constructor.
    /// </summary>
    public RtfColorTable()
    {
        colors = new List<int>();
    }

    /// <summary>
    /// Insert new color to the list.
    /// </summary>
    /// <param name="color">New color to insert.</param>
    public void AddColor(Color color)
    {
        colors.Add(color.ToArgb());
    }

    /// <summary>
    /// Get color at index.
    /// </summary>
    /// <param name="index">Index of the color.</param>
    /// <returns>Color at the index in the table.</returns>
    public Color this[int index] => Color.FromArgb(colors[index]);

    /// <summary>
    /// Number of colors in the table
    /// </summary>
    public int Count => colors.Count;

    /// <summary>
    /// Get the index of a color in the table.
    /// </summary>
    /// <param name="color">Color to get the index of.</param>
    /// <returns>Index of the color.</returns>
    public int IndexOf(Color color)
    {
        return colors.IndexOf(color.ToArgb());
    }
}
