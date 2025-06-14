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
 * Class:		InfoGroup
 * Description:	Class to encapsulate all the information contained in an RTF group of type "\info".
 * ******************************************************************************/

using System.Text;

namespace Net.Sgoliver.NRtfTree.Util;
/// <summary>
/// Class that encapsulates all the information contained in a "\info" group of an RTF document.
/// </summary>
public class InfoGroup
{

    #region Properties

    /// <summary>
    /// Title of the document.
    /// </summary>
    public string Title { get; set; } = "";

    /// <summary>
    /// Subject of the document.
    /// </summary>
    public string Subject { get; set; } = "";

    /// <summary>
    /// Author of the document.
    /// </summary>
    public string Author { get; set; } = "";

    /// <summary>
    /// Manager of the document.
    /// </summary>
    public string Manager { get; set; } = "";

    /// <summary>
    /// Company of the author of the document.
    /// </summary>
    public string Company { get; set; } = "";

    /// <summary>
    /// Last person who made changes to the document.
    /// </summary>
    public string Operator { get; set; } = "";

    /// <summary>
    /// Category of the document.
    /// </summary>
    public string Category { get; set; } = "";

    /// <summary>
    /// Keywords of the document.
    /// </summary>
    public string Keywords { get; set; } = "";

    /// <summary>
    /// Comments.
    /// </summary>
    public string Comment { get; set; } = "";

    /// <summary>
    /// Comments displayed in the "Summary Info" or "Properties" text box in Microsoft Word.
    /// </summary>
    public string DocComment { get; set; } = "";

    /// <summary>
    /// The base address used in relative paths for document links. This can be a local path or a URL.
    /// </summary>
    public string HlinkBase { get; set; } = "";

    /// <summary>
    /// Date/Time the document was created.
    /// </summary>
    public DateTime CreationTime { get; set; } = DateTime.MinValue;

    /// <summary>
    /// Date/Time of document review.
    /// </summary>
    public DateTime RevisionTime { get; set; } = DateTime.MinValue;

    /// <summary>
    /// Date/Time the document was last printed.
    /// </summary>
    public DateTime LastPrintTime { get; set; } = DateTime.MinValue;

    /// <summary>
    /// Date/Time of last copy of the document.
    /// </summary>
    public DateTime BackupTime { get; set; } = DateTime.MinValue;

    /// <summary>
    /// Version of the document.
    /// </summary>
    public int Version { get; set; } = -1;

    /// <summary>
    /// Internal version of the document.
    /// </summary>
    public int InternalVersion { get; set; } = -1;

    /// <summary>
    /// Total time spent editing the document (in minutes).
    /// </summary>
    public int EditingTime { get; set; } = -1;

    /// <summary>
    /// Number of pages in the document.
    /// </summary>
    public int NumberOfPages { get; set; } = -1;

    /// <summary>
    /// Number of words in the document.
    /// </summary>
    public int NumberOfWords { get; set; } = -1;

    /// <summary>
    /// Number of characters in the document.
    /// </summary>
    public int NumberOfChars { get; set; } = -1;

    /// <summary>
    /// Internal identification of the document.
    /// </summary>
    public int Id { get; set; } = -1;

    #endregion

    #region Public Methods

    /// <summary>
    /// Returns the representation of the node as a string.
    /// </summary>
    /// <returns>Representation of the node in the form of a character string.</returns>
    public override string ToString()
    {
        var str = new StringBuilder();

        str.AppendLine("Title     : " + Title);
        str.AppendLine("Subject   : " + Subject);
        str.AppendLine("Author    : " + Author);
        str.AppendLine("Manager   : " + Manager);
        str.AppendLine("Company   : " + Company);
        str.AppendLine("Operator  : " + Operator);
        str.AppendLine("Category  : " + Category);
        str.AppendLine("Keywords  : " + Keywords);
        str.AppendLine("Comment   : " + Comment);
        str.AppendLine("DComment  : " + DocComment);
        str.AppendLine("HLinkBase : " + HlinkBase);
        str.AppendLine("Created   : " + CreationTime);
        str.AppendLine("Revised   : " + RevisionTime);
        str.AppendLine("Printed   : " + LastPrintTime);
        str.AppendLine("Backup    : " + BackupTime);
        str.AppendLine("Version   : " + Version);
        str.AppendLine("IVersion  : " + InternalVersion);
        str.AppendLine("Editing   : " + EditingTime);
        str.AppendLine("Num Pages : " + NumberOfPages);
        str.AppendLine("Num Words : " + NumberOfWords);
        str.AppendLine("Num Chars : " + NumberOfChars);
        str.AppendLine("Id        : " + Id);

        return str.ToString();
    }

    #endregion
}
