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
 * Class:		RtfTree
 * Description:	Representa un documento RTF en forma de �rbol.
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Util;
using SkiaSharp;

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Representation of a tree -shaped RTF document.
/// </summary>
public class RtfTree
{
    #region Private fields

    /// <summary>
    /// Rtf root node.
    /// </summary>
    private RtfTreeNode rootNode;
    /// <summary>
    /// RTF input file/stream
    /// </summary>
    private TextReader? rtf;
    /// <summary>
    /// Lexical analyzer for RTF
    /// </summary>
    private RtfLex? lex;
    /// <summary>
    /// RTF Token
    /// </summary>
    private RtfToken? tok;
    /// <summary>
    /// Current node depth
    /// </summary>
    private int level;
    /// <summary>
    /// Indicate whether the special characters (\') are decoded by joining contiguous text nodes.
    /// </summary>
    private bool mergeSpecialCharacters;

    #endregion

    #region Constructors

    /// <summary>
    /// RtfTree constructor.
    /// </summary>
    public RtfTree()
    {
        //The root node of the document is created
        rootNode = new RtfTreeNode(RtfNodeType.Root,"ROOT",false,0);

        rootNode.Tree = this;

        /* Initialized default */

        mergeSpecialCharacters = false;
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Make an exact copy of the RTF tree.
    /// </summary>
    /// <returns>Returns an exact copy of the RTF tree.</returns>
    public RtfTree CloneTree()
    {
        var clon = new RtfTree();

        clon.rootNode = rootNode.CloneNode();

        return clon;
    }

    /// <summary>
    /// Load a file in RTF format.
    /// </summary>
    /// <param name="path">Route of the file with the document.</param>
    /// <returns>Value 0 is returned in case of no error in the document loading.
    /// incase of error the value -1.</returns>
    public int LoadRtfFile(string path)
    {
        var res = 0;

        rtf = new StreamReader(path);

        lex = new RtfLex(rtf);

        res = parseRtfTree();

        rtf.Close();

        return res;
    }

    /// <summary>
    /// Load a text stream with RTF format.
    /// </summary>
    /// <param name="text">Text stream containing the document.</param>
    /// <returns>Value 0 is returned in case of no error in the document loading.
    /// incase of error the value -1.</returns>
    public int LoadRtfText(string text)
    {
        var res = 0;

        rtf = new StringReader(text);

        lex = new RtfLex(rtf);

        res = parseRtfTree();

        rtf.Close();

        return res;
    }

    /// <summary>
    /// Write the RTF code of the document to a file.
    /// </summary>
    /// <param name="filePath">Route of the file to generate from the RTF document.</param>
    public void SaveRtf(string filePath)
    {
        var sw = new StreamWriter(filePath);

        //The RTF tree is transformed into text and written to file
        sw.Write(RootNode.Rtf);

        sw.Flush();
        sw.Close();
    }

    /// <summary>
    /// Returns a textual representation of the document loaded.
    /// </summary>
    /// <returns>String with the representation of the document.</returns>
    public override string ToString()
    {
        var res = "";

        res = toStringInm(rootNode, 0, false);

        return res;
    }

    /// <summary>
    /// Returns a textual representation of the document loaded. Add the node type to the left of the node content.
    /// </summary>
    /// <returns>String with the representation of the document.</returns>
    public string ToStringEx()
    {
        var res = "";

        res = toStringInm(rootNode, 0, true);

        return res;
    }

    /// <summary>
    /// Returns the RTF document font table.
    /// </summary>
    /// <returns>Table of fonts in the RTF document</returns>
    public RtfFontTable GetFontTable()
    {
        var tablaFuentes = new RtfFontTable();

        var root = rootNode;

        // Main Document Group
        var nprin = root.FirstChild;

        // We are looking for the font table in the tree
        var enc = false;
        var i = 0;
        var ntf = new RtfTreeNode();  //Node with the Table of Fonts

        while (!enc && nprin is { ChildNodes: not null } && i < nprin.ChildNodes.Count)
        {
            if (nprin.ChildNodes[i].NodeType == RtfNodeType.Group &&
                nprin.ChildNodes[i].FirstChild?.NodeKey == "fonttbl")
            {
                enc = true;
                ntf = nprin.ChildNodes[i];
            }

            i++;
        }

        // We fill the sources array
        if (ntf.ChildNodes == null) return tablaFuentes;
        
        for (var j = 1; j < ntf.ChildNodes.Count; j++)
        {
            var fuente = ntf.ChildNodes[j];

            var indiceFuente = -1;
            string? nombreFuente = null;

            if (fuente.ChildNodes != null)
                foreach (RtfTreeNode nodo in fuente.ChildNodes)
                {
                    if (nodo.NodeKey == "f")
                        indiceFuente = nodo.Parameter;

                    if (nodo.NodeType == RtfNodeType.Text)
                        nombreFuente = nodo.NodeKey[..^1];
                }

            if (nombreFuente != null) 
                tablaFuentes.AddFont(indiceFuente, nombreFuente);
        }

        return tablaFuentes;
    }

    /// <summary>
    /// Returns the color table of the RTF document.
    /// </summary>
    /// <returns>RTF document colors table</returns>
    public RtfColorTable GetColorTable()
    {
        var tableOfColors = new RtfColorTable();

        var root = rootNode;

        // Main Document Group
        var nprin = root.FirstChild;

        // We are looking for the color table in the tree
        var enc = false;
        var i = 0;
        var ntc = new RtfTreeNode();

        while (!enc && nprin is { ChildNodes: not null } && i < nprin.ChildNodes.Count)
        {
            if (nprin.ChildNodes[i]?.NodeType == RtfNodeType.Group &&
                nprin.ChildNodes[i]?.FirstChild?.NodeKey == "colortbl")
            {
                enc = true;
                ntc = nprin.ChildNodes[i];
            }

            i++;
        }

        // We fill the array of colors
        var red = 0;
        var green = 0;
        var blue = 0;

        //We add the default color, in this case the black.
        //tabla.Add(Color.FromArgb(rojo,verde,azul));

        if (ntc?.ChildNodes == null) return tableOfColors;
        
        for (var j = 1; j < ntc.ChildNodes.Count; j++)
        {
            var node = ntc.ChildNodes[j];
            if (node == null) continue;
            
            if (node.NodeType == RtfNodeType.Text && node.NodeKey.Trim() == ";")
            {
                tableOfColors.AddColor(new SKColor((byte)red, (byte)green, (byte)blue));

                red = 0;
                green = 0;
                blue = 0;
            }
            else if (node.NodeType == RtfNodeType.Keyword)
            {
                switch (node.NodeKey)
                {
                    case "red":
                        red = node.Parameter;
                        break;
                    case "green":
                        green = node.Parameter;
                        break;
                    case "blue":
                        blue = node.Parameter;
                        break;
                }
            }
        }

        return tableOfColors;
    }

    /// <summary>
    /// Returns the RTF document style sheets.
    /// </summary>
    /// <returns>Table of stylesheets in the RTF document.</returns>
    public RtfStyleSheetTable GetStyleSheetTable()
    {
        var sstable = new RtfStyleSheetTable();

        var sst = MainGroup?.SelectSingleGroup("stylesheet");

        var styles = sst?.ChildNodes;

        if (styles == null) return sstable;
        
        for (var i = 1; i < styles.Count; i++)
        {
            var style = styles[i];

            var rtfss = ParseStyleSheet(style);

            sstable.AddStyleSheet(rtfss.Index, rtfss);
        }

        return sstable;
    }

    /// <summary>
    /// Returns the information contained in the "\info" group of the RTF document.
    /// </summary>
    /// <returns>InfoGroup object containing the info from the document.</returns>
    public InfoGroup? GetInfoGroup()
    {
        InfoGroup? info = null;

        var infoNode = RootNode.SelectSingleNode("info");

        // If there is the node "\info" we will extract all the information.
        if (infoNode == null) return info;
        
        RtfTreeNode? auxnode = null;

        info = new InfoGroup();

        //Title
        if ((auxnode = rootNode.SelectSingleNode("title")) != null)
            info.Title = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Subject
        if ((auxnode = rootNode.SelectSingleNode("subject")) != null)
            info.Subject = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Author
        if ((auxnode = rootNode.SelectSingleNode("author")) != null)
            info.Author = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Manager
        if ((auxnode = rootNode.SelectSingleNode("manager")) != null)
            info.Manager = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Company
        if ((auxnode = rootNode.SelectSingleNode("company")) != null)
            info.Company = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Operator
        if ((auxnode = rootNode.SelectSingleNode("operator")) != null)
            info.Operator = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Category
        if ((auxnode = rootNode.SelectSingleNode("category")) != null)
            info.Category = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Keywords
        if ((auxnode = rootNode.SelectSingleNode("keywords")) != null)
            info.Keywords = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Comments
        if ((auxnode = rootNode.SelectSingleNode("comment")) != null)
            info.Comment = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Document comments
        if ((auxnode = rootNode.SelectSingleNode("doccomm")) != null)
            info.DocComment = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Hlinkbase (The base address that is used for the path of all relative hyperlinks inserted in the document)
        if ((auxnode = rootNode.SelectSingleNode("hlinkbase")) != null)
            info.HlinkBase = auxnode.NextSibling?.NodeKey ?? string.Empty;

        //Version
        if ((auxnode = rootNode.SelectSingleNode("version")) != null)
            info.Version = auxnode.Parameter;

        //Internal Version
        if ((auxnode = rootNode.SelectSingleNode("vern")) != null)
            info.InternalVersion = auxnode.Parameter;

        //Editing Time
        if ((auxnode = rootNode.SelectSingleNode("edmins")) != null)
            info.EditingTime = auxnode.Parameter;

        //Number of Pages
        if ((auxnode = rootNode.SelectSingleNode("nofpages")) != null)
            info.NumberOfPages = auxnode.Parameter;

        //Number of Chars
        if ((auxnode = rootNode.SelectSingleNode("nofchars")) != null)
            info.NumberOfChars = auxnode.Parameter;

        //Number of Words
        if ((auxnode = rootNode.SelectSingleNode("nofwords")) != null)
            info.NumberOfWords = auxnode.Parameter;

        //Id
        if ((auxnode = rootNode.SelectSingleNode("id")) != null)
            info.Id = auxnode.Parameter;

        //Creation DateTime
        if ((auxnode = rootNode.SelectSingleNode("creatim")) != null)
            if (auxnode.ParentNode != null)
                info.CreationTime = parseDateTime(auxnode.ParentNode);

        //Revision DateTime
        if ((auxnode = rootNode.SelectSingleNode("revtim")) != null)
            if (auxnode.ParentNode != null)
                info.RevisionTime = parseDateTime(auxnode.ParentNode);

        //Last Print Time
        if ((auxnode = rootNode.SelectSingleNode("printim")) != null)
            if (auxnode.ParentNode != null)
                info.LastPrintTime = parseDateTime(auxnode.ParentNode);

        //Backup Time
        if ((auxnode = rootNode.SelectSingleNode("buptim")) != null)
            if (auxnode.ParentNode != null)
                info.BackupTime = parseDateTime(auxnode.ParentNode);

        return info;
    }

    /// <summary>
    /// Returns the code table with which the RTF document is coded.
    /// </summary>
    /// <returns>RTF document code table. If not specified in the document, the current system code table is returned.4</returns>
    public Encoding GetEncoding()
    {
        //Contributed by Jan Stuchlík

        var encoding = Encoding.Default;

        var cpNode = RootNode.SelectSingleNode("ansicpg");

        if (cpNode != null)
        {
            encoding = Encoding.GetEncoding(cpNode.Parameter);
        }

        return encoding;
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Analyze the document and load its tree structure.
    /// </summary>
    /// <returns>Value 0 is returned in case of no error in the document loading.
    /// incase of error the value -1.</returns>
    private int parseRtfTree()
    {
        var res = 0;

        // Document default coding
        var encoding = Encoding.Default;

        var curNode = rootNode;

        // New nodes to build the RTF tree

        // The first token is obtained
        tok = lex?.NextToken();

        while (tok != null && tok.Type != RtfTokenType.Eof)
        {
            RtfTreeNode? newNode;
            switch (tok.Type)
            {
                case RtfTokenType.GroupStart:
                    newNode = new RtfTreeNode(RtfNodeType.Group,"GROUP",false,0);
                    curNode?.AppendChild(newNode);
                    curNode = newNode;
                    level++;
                    break;
                case RtfTokenType.GroupEnd:
                    curNode = curNode?.ParentNode;
                    level--;
                    break;
                case RtfTokenType.Keyword:
                case RtfTokenType.Control:
                case RtfTokenType.Text:
                    if (mergeSpecialCharacters)
                    {
                        //Contributed by Jan Stuchlík
                        var isText = tok.Type == RtfTokenType.Text || (tok.Type == RtfTokenType.Control && tok.Key == "'");
                        if (curNode?.LastChild is { NodeType: RtfNodeType.Text } && isText)
                        {
                            if (tok.Type == RtfTokenType.Text)
                            {
                                curNode.LastChild.NodeKey += tok.Key;
                                break;
                            }
                            if (tok.Type == RtfTokenType.Control && tok.Key == "'")
                            {
                                curNode.LastChild.NodeKey += DecodeControlChar(tok.Parameter, encoding);
                                break;
                            }
                        }
                        else
                        {
                            // First special character \'
                            if (tok.Type == RtfTokenType.Control && tok.Key == "'")
                            {
                                newNode = new RtfTreeNode(RtfNodeType.Text, DecodeControlChar(tok.Parameter, encoding), false, 0);
                                curNode?.AppendChild(newNode);
                                break;
                            }
                        }
                    }

                    newNode = new RtfTreeNode(tok);
                    curNode?.AppendChild(newNode);

                    if (mergeSpecialCharacters)
                    {
                        //Contributed by Jan Stuchlík
                        if (level == 1 && newNode.NodeType == RtfNodeType.Keyword && newNode.NodeKey == "ansicpg")
                        {
                            encoding = Encoding.GetEncoding(newNode.Parameter);
                        }
                    }

                    break;
                default:
                    res = -1;
                    break;
            }

            // The following token is obtained
            tok = lex?.NextToken();
        }

        // If the current level is not 0 (== a group is not well formed)
        if (level != 0)
        {
            res = -1;
        }

        return res;
    }

    /// <summary>
    /// Decodes a special character indicated by its decimal code
    /// </summary>
    /// <param name="code">Special character code (\')</param>
    /// <param name="enc">Coding used to decode the special character.</param>
    /// <returns>Special Decoded character.</returns>
    private string DecodeControlChar(int code, Encoding enc)
    {
        //Contributed by Jan Stuchlík
        return enc.GetString(new[] {(byte)code});                
    }

    /// <summary>
    /// Auxiliary method to generate the textual representation of the RTF document.
    /// </summary>
    /// <param name="curNode">Current tree node.</param>
    /// <param name="levelParam">Current level in tree.</param>
    /// <param name="showNodeTypes">Indicates whether the type of each tree node will be displayed.</param>
    /// <returns>Textual representation of the 'curNode' node with level 'level'</returns>
    private string toStringInm(RtfTreeNode curNode, int levelParam, bool showNodeTypes)
    {
        var res = new StringBuilder();

        var children = curNode.ChildNodes;

        for (var i = 0; i < levelParam; i++)
            res.Append("  ");

        if (curNode.NodeType == RtfNodeType.Root)
            res.Append("ROOT\r\n");
        else if (curNode.NodeType == RtfNodeType.Group)
            res.Append("GROUP\r\n");
        else
        {
            if (showNodeTypes)
            {
                res.Append(curNode.NodeType);
                res.Append(": ");
            }

            res.Append(curNode.NodeKey);

            if (curNode.HasParameter)
            {
                res.Append(' ');
                res.Append(Convert.ToString(curNode.Parameter));
            }

            res.Append("\r\n");
        }

        if (children == null) return res.ToString();
        
        foreach (RtfTreeNode node in children)
        {
            res.Append(toStringInm(node, levelParam + 1, showNodeTypes));
        }

        return res.ToString();
    }

    /// <summary>
    /// Parse a format date "\yr2005\mo12\dy2\hr22\min56\sec15"
    /// </summary>
    /// <param name="group">RTF Group with the date.</param>
    /// <returns>Parsed DateTime object.</returns>
    private static DateTime parseDateTime(RtfTreeNode group)
    {
        int year = 0, month = 0, day = 0, hour = 0, min = 0, sec = 0;

        if (group.ChildNodes != null)
            foreach (RtfTreeNode node in group.ChildNodes)
            {
                switch (node.NodeKey)
                {
                    case "yr":
                        year = node.Parameter;
                        break;
                    case "mo":
                        month = node.Parameter;
                        break;
                    case "dy":
                        day = node.Parameter;
                        break;
                    case "hr":
                        hour = node.Parameter;
                        break;
                    case "min":
                        min = node.Parameter;
                        break;
                    case "sec":
                        sec = node.Parameter;
                        break;
                }
            }

        var dt = new DateTime(year, month, day, hour, min, sec);

        return dt;
    }

    /// <summary>
    /// Extract the text from an RTF tree.
    /// </summary>
    /// <returns>Plaintext of the document.</returns>
    private string ConvertToText()
    {
        var res = new StringBuilder("");

        var pardNode =
            MainGroup?.SelectSingleChildNode("pard");

        if (pardNode == null) return res.ToString();

        if (MainGroup?.ChildNodes == null) return res.ToString();
        
        for (var i = pardNode.Index; i < MainGroup.ChildNodes.Count; i++)
        {
            res.Append(MainGroup.ChildNodes[i]?.Text);
        }

        return res.ToString();
    }

    /// <summary>
    /// Parse a node as a document stylesheet.
    /// </summary>
    /// <param name="ssnode">Node with the styles.</param>
    /// <returns>Table of styles for the document.</returns>
    private RtfStyleSheet ParseStyleSheet(RtfTreeNode ssnode)
    {
        var rss = new RtfStyleSheet();

        if (ssnode.ChildNodes == null) return rss;
        
        foreach (RtfTreeNode node in ssnode.ChildNodes)
        {
            switch (node.NodeKey)
            {
                case "cs":
                    rss.Type = RtfStyleSheetType.Character;
                    rss.Index = node.Parameter;
                    break;
                case "s":
                    rss.Type = RtfStyleSheetType.Paragraph;
                    rss.Index = node.Parameter;
                    break;
                case "ds":
                    rss.Type = RtfStyleSheetType.Section;
                    rss.Index = node.Parameter;
                    break;
                case "ts":
                    rss.Type = RtfStyleSheetType.Table;
                    rss.Index = node.Parameter;
                    break;
                case "additive":
                    rss.Additive = true;
                    break;
                case "sbasedon":
                    rss.BasedOn = node.Parameter;
                    break;
                case "snext":
                    rss.Next = node.Parameter;
                    break;
                case "sautoupd":
                    rss.AutoUpdate = true;
                    break;
                case "shidden":
                    rss.Hidden = true;
                    break;
                case "slink":
                    rss.Link = node.Parameter;
                    break;
                case "slocked":
                    rss.Locked = true;
                    break;
                case "spersonal":
                    rss.Personal = true;
                    break;
                case "scompose":
                    rss.Compose = true;
                    break;
                case "sreply":
                    rss.Reply = true;
                    break;
                case "styrsid":
                    rss.Styrsid = node.Parameter;
                    break;
                case "ssemihidden":
                    rss.SemiHidden = true;
                    break;
                default:
                {
                    if (node.NodeType == RtfNodeType.Group &&
                        (node.ChildNodes?[0]?.NodeKey == "*" && node.ChildNodes[1]?.NodeKey == "keycode"))
                    {
                        rss.KeyCode = new RtfNodeCollection();

                        for (var i = 2; i < node.ChildNodes.Count; i++)
                        {
                            rss.KeyCode.Add(node.ChildNodes[i]?.CloneNode());
                        }
                    }
                    else if (node.NodeType == RtfNodeType.Text)
                    {
                        rss.Name = node.NodeKey[..^1];
                    }
                    else
                    {
                        if (node.NodeKey != "*")
                            rss.Formatting?.Add(node);
                    }

                    break;
                }
            }
        }

        return rss;
    }

    #endregion

    #region Properties

    /// <summary>
    /// Return the root node of the document tree.
    /// </summary>
    public RtfTreeNode RootNode => rootNode;

    /// <summary>
    /// Returns the main group of the document.
    /// </summary>
    public RtfTreeNode? MainGroup =>
        // The main group (null incase it doesnt exist)
        rootNode.HasChildNodes() ? rootNode.ChildNodes?[0] : null;

    /// <summary>
    /// Returns the document as RTF.
    /// </summary>
    public string Rtf => rootNode.Rtf;

    /// <summary>
    /// Indicate whether the special characters (\') are decoded by joining contiguous text nodes.
    /// </summary>
    public bool MergeSpecialCharacters
    {
        get => mergeSpecialCharacters;
        set => mergeSpecialCharacters = value;
    }

    /// <summary>
    /// Return the document as plaintext
    /// </summary>
    public string Text => ConvertToText();

    #endregion
}
