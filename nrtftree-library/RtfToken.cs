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
 * Class:       RtfToken
 * Description: Token read by the lexical analyzer for RTF documents.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Token read by the lexical analyzer for RTF documents.
/// </summary>
public class RtfToken
{

    /// <summary>
    /// Token type.
    /// </summary>
    public RtfTokenType Type { get; set; }

    /// <summary>
    /// Keyword / Control symbol / character.
    /// </summary>
    public string Key { get; set; }

    /// <summary>
    /// Indicate if the token has an associated parameter.
    /// </summary>
    public bool HasParameter { get; set; }

    /// <summary>
    /// Keyword parameter or control symbol.
    /// </summary>
    public int Parameter { get; set; }


    /// <summary>
    /// RtfToken class builder. Create an empty token.
    /// </summary>
    public RtfToken()
    {
        Type = RtfTokenType.None;
        Key = "";
    }

}

