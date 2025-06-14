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
 * Class:		RtfMerger
 * Description:	Class to combine multiple RTF documents.
 * Notes:       Originally contributed by Fabio Borghi.
 * ******************************************************************************/

using Net.Sgoliver.NRtfTree.Util;
using System.Drawing;

namespace Net.Sgoliver.NRtfTree.Core;

/// <summary>
/// Class to combine multiple RTF documents.
/// </summary>
public class RtfMerger
{
    private bool removeLastPar;

    #region Constructors

    /// <summary>
    /// Constructor. 
    /// </summary>
    /// <param name="templatePath">Template document path.</param>
    public RtfMerger(string templatePath)
    {
        // The source document is loaded
        Template = new RtfTree();
        Template.LoadRtfFile(templatePath);

        // The list of replacement parameters (placeholders) is created
        Placeholders = new Dictionary<string, RtfTree>();
    }

    /// <summary>
    /// Constructor. 
    /// </summary>
    /// <param name="templateTree">Template document path.</param>
    public RtfMerger(RtfTree templateTree)
    {
        // The source document is loaded
        Template = templateTree;
        
        // The list of replacement parameters (placeholders) is created
        Placeholders = new Dictionary<string, RtfTree>();
    }

    /// <summary>
    /// Constructor. 
    /// </summary>
    public RtfMerger()
    {
        // The list of replacement parameters (placeholders) is created
        Placeholders = new Dictionary<string, RtfTree>();
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Associates a new placeholder with the path of the document to be inserted.
    /// </summary>
    /// <param name="ph">Name of the placeholder.</param>
    /// <param name="path">Path of the document to be inserted.</param>
    public void AddPlaceHolder(string ph, string path)
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(path);

        if (res == 0)
        {
            Placeholders.Add(ph, tree);
        }
    }

    /// <summary>
    /// Associates a new placeholder with the path of the document to be inserted.
    /// </summary>
    /// <param name="ph">Name of the placeholder.</param>
    /// <param name="docTree">RTF tree of the document to be inserted.</param>
    public void AddPlaceHolder(string ph, RtfTree docTree)
    {
        Placeholders.Add(ph, docTree);
    }

    /// <summary>
    /// Disassociates a placeholder parameter with the path of the document to be inserted.
    /// </summary>
    /// <param name="ph">Name of the placeholder.</param>
    public void RemovePlaceHolder(string ph)
    {
        Placeholders.Remove(ph);
    }

    /// <summary>
    /// Performs the combination of RTF documents.
    /// <param name="removeLastParagraph">Indicates whether the last \par node should be removed from documents inserted into the template.</param>
    /// <returns>Returns the RTF tree resulting from the merge.</returns>
    /// </summary>
    public RtfTree? Merge(bool removeLastParagraph = true)
    {
        this.removeLastPar = removeLastParagraph;

        // The main group of the tree is obtained
        var parentNode = Template?.MainGroup;

        // If the document has a main group
        if (parentNode != null)
        {
            // The document text is analyzed for replacement parameters and the documents are combined
            analyzeTextContent(parentNode);
        }

        return Template;
    }

    #endregion

    #region Properties

    /// <summary>
    /// Returns the list of replacement parameters in the format: [string, RtfTree]
    /// </summary>
    public Dictionary<string, RtfTree> Placeholders { get; }

    /// <summary>
    /// Gets or sets the RTF tree of the template document.
    /// </summary>
    public RtfTree? Template { get; set; }

    #endregion

    #region Private Methods

    /// <summary>
    /// Analyzes the document text for replacement parameters and merges the documents.
    /// </summary>
    /// <param name="parentNode">Tree node to be processed.</param>
    private void analyzeTextContent(RtfTreeNode? parentNode)
    {
        int indPH;

        // If the node is of type group and contains child nodes
        if (parentNode == null || !parentNode.HasChildNodes()) return;
        if (parentNode.ChildNodes == null) return;
        
        //All child nodes are traversed
        for (var iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
        {
            var currNode = parentNode.ChildNodes[iNdIndex];

            // If the current node is of type Text, labels to replace are searched for.
            if (currNode is { NodeType: RtfNodeType.Text })
            {
                // All configured tags are traversed
                foreach (var ph in Placeholders.Keys)
                {
                    // Search for the next occurrence of the current tag
                    indPH = currNode.NodeKey.IndexOf(ph, StringComparison.Ordinal);

                    // If a tag has been found
                    if (indPH == -1) continue;
                    // The tree to be inserted into the current tag is retrieved
                    var docToInsert = Placeholders[ph].CloneTree();

                    // The new tree is inserted into the base tree
                    mergeCore(parentNode, iNdIndex, docToInsert, ph, indPH);

                    // Since the current node may have changed, we decrement the index.
                    // and we exit the loop to analyze it again
                    iNdIndex--;
                    break;
                }
            }
            else
            {
                // If the current node has children, the child nodes are analyzed.
                if (currNode != null && currNode.HasChildNodes())
                {
                    analyzeTextContent(currNode);
                }
            }
        }
    }

    /// <summary>
    /// Inserts a new tree in place of a text label in the base tree.
    /// </summary>
    /// <param name="parentNode">Group type node being processed.</param>
    /// <param name="iNdIndex">Index (within the parent group) of the text node being processed</param>
    /// <param name="docToInsert">New RTF tree to insert.</param>
    /// <param name="strCompletePlaceholder">Text of the label to be replaced.</param>
    /// <param name="intPlaceHolderNodePos">Position of the label to be replaced within the text node being processed</param>
    private void mergeCore(RtfTreeNode? parentNode, int iNdIndex, RtfTree docToInsert, string strCompletePlaceholder, int intPlaceHolderNodePos)
    {
        // If the document to be inserted is not empty
        if (!docToInsert.RootNode.HasChildNodes()) return;
        
        var currentIndex = iNdIndex + 1;

        // The color tables are combined and the colors of the document to be inserted are adjusted.
        mainAdjustColor(docToInsert);

        // The font tables are combined and the fonts of the document to be inserted are adjusted.
        mainAdjustFont(docToInsert);

        // The header information of the document to be inserted is removed (colors, fonts, info, ...)
        cleanToInsertDoc(docToInsert);

        // If the document to be inserted has content
        if (docToInsert.RootNode.FirstChild != null && docToInsert.RootNode.FirstChild.HasChildNodes())
        {
            // The new document is inserted into the base tree
            execMergeDoc(parentNode, docToInsert, currentIndex);
        }

        // If the label is not at the end of the text node:
        // A text node is inserted with the rest of the original text (removing the label)
        if (parentNode != null && parentNode.ChildNodes?[iNdIndex]?.NodeKey.Length != (intPlaceHolderNodePos + strCompletePlaceholder.Length))
        {
            //A text node is inserted with the rest of the original text (removing the label)
            var remText = 
                (parentNode.ChildNodes?[iNdIndex]
                    ?.NodeKey)?[(parentNode.ChildNodes[iNdIndex]!.NodeKey.IndexOf(strCompletePlaceholder, StringComparison.Ordinal) + strCompletePlaceholder.Length)..];
            
            if (remText == null) throw new NullReferenceException();
            parentNode.InsertChild(currentIndex + 1, new RtfTreeNode(RtfNodeType.Text, remText, false, 0));
        }

        // If the replaced tag was at the beginning of the text node we delete the
        // original node because it is no longer necessary
        if (intPlaceHolderNodePos == 0)
        {
            parentNode?.RemoveChild(iNdIndex);
        }
        // Otherwise we replace it with the text before the label
        else
        {
            if (parentNode == null) return;
            if (parentNode.ChildNodes != null)
                parentNode.ChildNodes[iNdIndex]!.NodeKey =
                    (parentNode.ChildNodes[iNdIndex]?.NodeKey)?[..intPlaceHolderNodePos] ?? throw new InvalidOperationException();
        }
    }

    /// <summary>
    /// Obtains the source code passed as a parameter, inserting it into the source table if necessary.
    /// </summary>
    /// <param name="fontDestTbl">Resulting font table.</param>
    /// <param name="sFontName">Font name.</param>
    /// <returns></returns>
    private int getFontID(ref RtfFontTable fontDestTbl, string sFontName)
    {
        var iExistingFontID = -1;

        if ((iExistingFontID = fontDestTbl.IndexOf(sFontName)) != -1) return iExistingFontID;

        fontDestTbl.AddFont(sFontName);
        iExistingFontID = fontDestTbl.IndexOf(sFontName);

        var nodeListToInsert = Template?.RootNode.SelectNodes("fonttbl");

        var ftFont = new RtfTreeNode(RtfNodeType.Group);
        ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, iExistingFontID));
        ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fnil", false, 0));
        ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Text, sFontName + ";", false, 0));
            
        nodeListToInsert?[0]?.ParentNode?.AppendChild(ftFont);

        return iExistingFontID;
    }

    /// <summary>
    /// Obtains the code of the color passed as a parameter, inserting it into the color table if necessary.
    /// </summary>
    /// <param name="colorDestTbl">Table of resulting colors.</param>
    /// <param name="iColorName">Color name.</param>
    /// <returns></returns>
    private int getColorID(RtfColorTable colorDestTbl, Color iColorName)
    {
        int iExistingColorID;

        if ((iExistingColorID = colorDestTbl.IndexOf(iColorName)) != -1) return iExistingColorID;
        
        iExistingColorID = colorDestTbl.Count;
        colorDestTbl.AddColor(iColorName);

        var nodeListToInsert = Template?.RootNode.SelectNodes("colortbl");

        nodeListToInsert?[0]?.ParentNode?.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "red", true, iColorName.R));
        nodeListToInsert?[0]?.ParentNode?.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "green", true, iColorName.G));
        nodeListToInsert?[0]?.ParentNode?.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "blue", true, iColorName.B));
        nodeListToInsert?[0]?.ParentNode?.AppendChild(new RtfTreeNode(RtfNodeType.Text, ";", false, 0));

        return iExistingColorID;
    }

    /// <summary>
    /// Adjusts the fonts of the document to be inserted.
    /// </summary>
    /// <param name="docToInsert">Document to insert.</param>
    private void mainAdjustFont(RtfTree docToInsert)
    {
        if(Template == null) return;
        var fontDestTbl = Template.GetFontTable();
        var fontToCopyTbl = docToInsert.GetFontTable();

        adjustFontRecursive(docToInsert.RootNode, fontDestTbl, fontToCopyTbl);
    }

    /// <summary>
    /// Adjusts the fonts of the document to be inserted.
    /// </summary>
    /// <param name="parentNode">Group node being processed.</param>
    /// <param name="fontDestTbl">Table of resulting fonts.</param>
    /// <param name="fontToCopyTbl">Table of fonts of the document to be inserted.</param>
    private void adjustFontRecursive(RtfTreeNode parentNode, RtfFontTable fontDestTbl, RtfFontTable fontToCopyTbl)
    {
        if (!parentNode.HasChildNodes()) return;
        if (parentNode.ChildNodes == null) return;
        
        for (var iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
        {
            if (parentNode.ChildNodes[iNdIndex] == null) continue;
            if (parentNode.ChildNodes[iNdIndex]!.NodeType == RtfNodeType.Keyword &&
                (parentNode.ChildNodes[iNdIndex]!.NodeKey == "f" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "stshfdbch" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "stshfloch" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "stshfhich" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "stshfbi" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "deff" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "af") &&
                parentNode.ChildNodes[iNdIndex]!.HasParameter)
            {
                parentNode.ChildNodes[iNdIndex]!.Parameter = getFontID(ref fontDestTbl,
                    fontToCopyTbl[parentNode.ChildNodes[iNdIndex]!.Parameter]);
            }

            adjustFontRecursive(parentNode.ChildNodes[iNdIndex]!, fontDestTbl, fontToCopyTbl);
        }
    }

    /// <summary>
    /// Adjusts the colors of the document to be inserted.
    /// </summary>
    /// <param name="docToInsert">Document to insert.</param>
    private void mainAdjustColor(RtfTree docToInsert)
    {
        if(Template == null) return;
        var colorDestTbl = Template.GetColorTable();
        var colorToCopyTbl = docToInsert.GetColorTable();

        adjustColorRecursive(docToInsert.RootNode, colorDestTbl, colorToCopyTbl);
    }

    /// <summary>
    /// Adjusts the colors of the document to be inserted.
    /// </summary>
    /// <param name="parentNode">Group node being processed.</param>
    /// <param name="colorDestTbl">Table of resulting colors.</param>
    /// <param name="colorToCopyTbl">Table of colors of the document to be inserted.</param>
    private void adjustColorRecursive(RtfTreeNode parentNode, RtfColorTable colorDestTbl, RtfColorTable colorToCopyTbl)
    {
        if (!parentNode.HasChildNodes()) return;
        if (parentNode.ChildNodes == null) return;

        for (var iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
        {
            if (parentNode.ChildNodes[iNdIndex] == null) continue;
            if (parentNode.ChildNodes[iNdIndex]!.NodeType == RtfNodeType.Keyword &&
                (parentNode.ChildNodes[iNdIndex]!.NodeKey == "cf" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "cb" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "pncf" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "brdrcf" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "cfpat" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "cbpat" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "clcfpatraw" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "clcbpatraw" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "ulc" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "chcfpat" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "chcbpat" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "highlight" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "clcbpat" ||
                 parentNode.ChildNodes[iNdIndex]!.NodeKey == "clcfpat") &&
                parentNode.ChildNodes[iNdIndex]!.HasParameter)
            {
                parentNode.ChildNodes[iNdIndex]!.Parameter = getColorID(colorDestTbl,
                    colorToCopyTbl[parentNode.ChildNodes[iNdIndex]!.Parameter]);
            }

            adjustColorRecursive(parentNode.ChildNodes[iNdIndex]!, colorDestTbl, colorToCopyTbl);
        }
    }

    /// <summary>
    /// Inserts the new tree into the base tree (as a new group) by removing all headers from the inserted document.
    /// </summary>
    /// <param name="parentNode">Base group in which the new tree will be inserted.</param>
    /// <param name="treeToCopyParent">New tree to insert.</param>
    /// <param name="intCurrIndex">Index at which the new tree will be inserted into the base group.</param>
    private void execMergeDoc(RtfTreeNode? parentNode, RtfTree treeToCopyParent, int intCurrIndex)
    {
        // Search for the first "\pard" in the document (beginning of text)
        var nodePard = treeToCopyParent.RootNode.FirstChild?.SelectSingleChildNode("pard");
        if(nodePard == null) return;

        // The index of the node within the main node is obtained
        if (treeToCopyParent.RootNode.FirstChild?.ChildNodes == null) return;
        
        var indPard = treeToCopyParent.RootNode.FirstChild.ChildNodes.IndexOf(nodePard);

        // The new group is created
        var newGroup = new RtfTreeNode(RtfNodeType.Group);

        // Paragraph and font options are reset
        newGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
        newGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));

        //Each child node of the new document is inserted into the base document
        for (var i = indPard + 1; i < treeToCopyParent.RootNode.FirstChild.ChildNodes.Count; i++)
        {
            var newNode = 
                treeToCopyParent.RootNode.FirstChild.ChildNodes[i]?.CloneNode();

            newGroup.AppendChild(newNode);
        }

        // The new group is inserted with the new document
        parentNode?.InsertChild(intCurrIndex, newGroup);
    }

    /// <summary>
    /// Removes unwanted elements from the document to be inserted, for example trailing "\par" nodes.
    /// </summary>
    /// <param name="docToInsert">Document to insert.</param>
    private void cleanToInsertDoc(RtfTree docToInsert)
    {
        // Deletes the last "\par" from the document if it exists
        var lastNode = docToInsert.RootNode.FirstChild?.LastChild;

        if (!removeLastPar) return;
        if (lastNode is { NodeType: RtfNodeType.Keyword, NodeKey: "par" })
        {
            docToInsert.RootNode.FirstChild?.RemoveChild(lastNode);
        }
    }

    #endregion
}

