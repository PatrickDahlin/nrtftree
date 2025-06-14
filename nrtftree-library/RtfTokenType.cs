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
 * Library:     NRtfTree
 * Version:     v0.4
 * Date:        29/06/2013
 * Copyright:   2006-2013 Salvador Gomez
 * Home Page:   http://www.sgoliver.net
 * GitHub:      https://github.com/sgolivernet/nrtftree
 * Class:       RtfTokenType
 * Description: Token types of an RTF document tree.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Token types of an RTF document tree.
/// </summary>
public enum RtfTokenType
{
    /// <summary>
    /// Indication that the Token has only just been initialized.
    /// </summary>
    None = 0,
    /// <summary>
    /// Keyword without parameter.
    /// </summary>
    Keyword = 1,
    /// <summary>
    /// Control symbol without parameter.
    /// </summary>
    Control = 2,
    /// <summary>
    /// Document text.
    /// </summary>
    Text = 3,
    /// <summary>
    /// Marker for end of file.
    /// </summary>
    Eof = 4,
    /// <summary>
    ///	Start of group: '{'
    /// </summary>
    GroupStart = 5,
    /// <summary>
    /// End of group: '}'
    /// </summary>
    GroupEnd = 6
}
