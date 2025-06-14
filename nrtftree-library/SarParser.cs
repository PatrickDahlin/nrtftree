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
 * Class:		SarParser
 * Description:	Abstract processor used by the RTF Reader class.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// This class, used by RTFReader, contains all the necessary
/// methods to treat each of the types of elements present in an RTF
/// document. These methods will be automatically called during the
/// analysis of the RTF document made by the RTF Reader class.
/// </summary>
public abstract class SarParser
{
    /// <summary>
    /// This method is called a single time at the beginning of the analysis of the RTF document.
    /// </summary>
    public abstract void StartRtfDocument();
    /// <summary>
    /// This method is called a only time at the end of the RTF document analysis.
    /// </summary>
    public abstract void EndRtfDocument();
    /// <summary>
    /// This method is called every time an RTF group start key is read.
    /// </summary>
    public abstract void StartRtfGroup();
    /// <summary>
    /// This method is called every time an RTF Group end key is read.
    /// </summary>
    public abstract void EndRtfGroup();
    /// <summary>
    /// This method is called every time an RTF keyword is read.
    /// </summary>
    /// <param name="key">Keyword read from the document.</param>
    /// <param name="hasParameter">Indicate if the keyword is accompanied by a parameter.</param>
    /// <param name="parameter">
    /// Parameter that accompanies the keyword. In the event that the
    /// keyword is not accompanied by any parameter, that is, the
    /// hasParam field is 'false', this field will contain 0 value.
    /// </param>
    public abstract void RtfKeyword(string key, bool hasParameter, int parameter);
    /// <summary>
    /// This method is called every time a RTF control symbol is read.
    /// </summary>
    /// <param name="key">Control symbol read of the document.</param>
    /// <param name="hasParameter">Indicate if the control symbol is accompanied by a parameter.</param>
    /// <param name="parameter">
    /// Parameter that accompanies the control symbol. In the event that
    /// the control symbol is not accompanied by any parameter, that is,
    /// the hasParam field is 'False', this field will contain the value 0.
    /// </param>
    public abstract void RtfControl(string key, bool hasParameter, int parameter);
    /// <summary>
    /// This method is called every time a text fragment of the RTF document is read.
    /// </summary>
    /// <param name="text">Text read from the document.</param>
    public abstract void RtfText(string text);
}
