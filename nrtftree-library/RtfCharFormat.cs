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
 * Class:		RtfCharFormat
 * Description:	Representation of formatted text.
 * ******************************************************************************/

using System.Drawing;

namespace Net.Sgoliver.NRtfTree.Util;

/// <summary>
/// Representation of formatted text.
/// </summary>
public class RtfCharFormat
{

    /// <summary>
    /// Bold font.
    /// </summary>
    public bool Bold { get; set; }

    /// <summary>
    /// Cursive font.
    /// </summary>
    public bool Italic { get; set; }

    /// <summary>
    /// Font underline.
    /// </summary>
    public bool Underline { get; set; }

    /// <summary>
    /// Font name.
    /// </summary>
    public string Font { get; set; } = "Arial";

    /// <summary>
    /// Font size.
    /// </summary>
    public int Size { get; set; } = 10;

    /// <summary>
    /// Font color.
    /// </summary>
    public Color Color { get; set; } = Color.Black;

}
