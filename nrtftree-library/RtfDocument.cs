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
 * Class:		RtfDocument
 * Description:	Class for generating RTF documents.
 * ******************************************************************************/

using System.Text;
using System.Globalization;
using Net.Sgoliver.NRtfTree.Core;
using SkiaSharp;

namespace Net.Sgoliver.NRtfTree.Util;

/// <summary>
/// Class for generating RTF documents.
/// </summary>
public class RtfDocument
{
    #region Private Fields

    /// <summary>
    /// Document encoding.
    /// </summary>
    private Encoding encoding;

    /// <summary>
    /// Table of fonts in the document.
    /// </summary>
    private RtfFontTable fontTable;

    /// <summary>
    /// Table of colors in the document
    /// </summary>
    private RtfColorTable colorTable;

    /// <summary>
    /// Main group of the document.
    /// </summary>
    private RtfTreeNode mainGroup;

    /// <summary>
    /// Current character format.
    /// </summary>
    private RtfCharFormat? currentFormat;

    /// <summary>
    /// Current paragraph format.
    /// </summary>
    private RtfParFormat currentParFormat;

    /// <summary>
    /// Document format.
    /// </summary>
    private RtfDocumentFormat docFormat;

    #endregion

    #region Constructor

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="enc">Document encoding.</param>
    public RtfDocument(Encoding enc)
    {
        encoding = enc;

        fontTable = new RtfFontTable();
        fontTable.AddFont("Arial");  //Default font

        colorTable = new RtfColorTable();
        colorTable.AddColor(SKColors.Black);  //Default color

        currentFormat = null;
        currentParFormat = new RtfParFormat();
        docFormat = new RtfDocumentFormat();

        mainGroup = new RtfTreeNode(RtfNodeType.Group);

        InitializeTree();
    }

    /// <summary>
    /// Constructor for the RtfDocument class. The default system encoding will be used.
    /// </summary>
    public RtfDocument() : this(Encoding.Default)
    {
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Save the document as an RTF file in the specified path.
    /// </summary>
    /// <param name="path">Path of the file to be created.</param>
    public void Save(string path)
    {
        var tree = GetTree();
        tree.SaveRtf(path);
    }

    /// <summary>
    /// Inserts a piece of text into the document with a specific text format.
    /// </summary>
    /// <param name="text">Text to insert.</param>
    /// <param name="format">Format of the text to insert.</param>
    public void AddText(string text, RtfCharFormat format)
    {
        UpdateFontTable(format);
        UpdateColorTable(format);

        UpdateCharFormat(format);

        InsertText(text);
    }

    /// <summary>
    /// Inserts a piece of text into the document using the current text formatting.
    /// </summary>
    /// <param name="text">Text to insert.</param>
    public void AddText(string text)
    {
        InsertText(text);
    }

    /// <summary>
    /// Inserts a specified number of line breaks into the document.
    /// </summary>
    /// <param name="n">Number of line breaks to insert.</param>
    public void AddNewLine(int n)
    {
        for(var i=0; i<n; i++)
            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
    }

    /// <summary>
    /// Inserts a line break in the document.
    /// </summary>
    public void AddNewLine()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
    }

    /// <summary>
    /// Start a new paragraph.
    /// </summary>
    public void AddNewParagraph()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));
    }

    /// <summary>
    /// Inserts a specified number of paragraph breaks into the document.
    /// </summary>
    /// <param name="n">Number of paragraph breaks to insert.</param>
    public void AddNewParagraph(int n)
    {
        for (var i = 0; i < n; i++)
            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));
    }

    /// <summary>
    /// Starts a new paragraph with the specified formatting.
    /// </summary>
    public void AddNewParagraph(RtfParFormat format)
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));

        UpdateParFormat(format);
    }

    /// <summary>
    /// Insert an image into the document.
    /// </summary>
    /// <param name="path">Path of the image to insert.</param>
    /// <param name="width">Desired width of the image in the document.</param>
    /// <param name="height">Desired height of the image in the document.</param>
    public void AddImage(string path, int width, int height)
    {
        FileStream fStream = null;
        BinaryReader br = null;

        try
        {
            byte[] data = null;

            var fInfo = new FileInfo(path);
            var numBytes = fInfo.Length;

            fStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            br = new BinaryReader(fStream);

            data = br.ReadBytes((int)numBytes);

            var hexdata = new StringBuilder();

            for (var i = 0; i < data.Length; i++)
            {
                hexdata.Append(Util.Text.GetHex(data[i]));
            }

            var imgdata = File.ReadAllBytes(path);
            var img = SKImage.FromEncodedData(imgdata);

            var imgGroup = new RtfTreeNode(RtfNodeType.Group);
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pict", false, 0));

            var format = "";
            if (path.ToLower().EndsWith("wmf"))
                format = "emfblip";
            else
                format = "jpegblip";

            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, format, false, 0));
            
            
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "picw", true, img.Width * 20));
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pich", true, img.Height * 20));
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "picwgoal", true, width * 20));
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pichgoal", true, height * 20));
            imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, hexdata.ToString(), false, 0));

            mainGroup.AppendChild(imgGroup);
        }
        finally
        {
            br?.Close();
            fStream?.Close();
        }
    }

    /// <summary>
    /// Sets the default character format.
    /// </summary>
    public void ResetCharFormat()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));
    }

    /// <summary>
    /// Sets the default paragraph format.
    /// </summary>
    public void ResetParFormat()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
    }

    /// <summary>
    /// Sets the default character and paragraph formatting.
    /// </summary>
    public void ResetFormat()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));
    }

    /// <summary>
    /// Updates the values of document format properties.
    /// </summary>
    /// <param name="format">Document format.</param>
    public void UpdateDocFormat(RtfDocumentFormat format)
    {
        docFormat.MarginL = format.MarginL;
        docFormat.MarginR = format.MarginR;
        docFormat.MarginT = format.MarginT;
        docFormat.MarginB = format.MarginB;
    }

    /// <summary>
    /// Updates the values of text and paragraph formatting properties.
    /// </summary>
    /// <param name="format">Format of text to be inserted.</param>
    public void UpdateCharFormat(RtfCharFormat format)
    {
        if (currentFormat != null)
        {
            SetFormatColor(format.Color);
            SetFormatSize(format.Size);
            SetFormatFont(format.Font);

            SetFormatBold(format.Bold);
            SetFormatItalic(format.Italic);
            SetFormatUnderline(format.Underline);
        }
        else //currentFormat == null
        {
            var indColor = colorTable.IndexOf(format.Color);

            if (indColor == -1)
            {
                colorTable.AddColor(format.Color);
                indColor = colorTable.IndexOf(format.Color);
            }

            var indFont = fontTable.IndexOf(format.Font);

            if (indFont == -1)
            {
                fontTable.AddFont(format.Font);
                indFont = fontTable.IndexOf(format.Font);
            }

            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "cf", true, indColor));
            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fs", true, format.Size * 2));
            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, indFont));

            if (format.Bold)
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "b", false, 0));

            if (format.Italic)
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "i", false, 0));

            if (format.Underline)
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ul", false, 0));

            currentFormat = new RtfCharFormat();
            currentFormat.Color = format.Color;
            currentFormat.Size = format.Size;
            currentFormat.Font = format.Font;
            currentFormat.Bold = format.Bold;
            currentFormat.Italic = format.Italic;
            currentFormat.Underline = format.Underline;
        }
    }

    /// <summary>
    /// Sets the paragraph format passed as a parameter.
    /// </summary>
    /// <param name="format">Paragraph format to use.</param>
    public void UpdateParFormat(RtfParFormat format)
    {
        SetAlignment(format.Alignment);
        SetLeftIndentation(format.LeftIndentation);
        SetRightIndentation(format.RightIndentation);
    }

    /// <summary>
    /// Set the alignment of the text within the paragraph.
    /// </summary>
    /// <param name="align">Alignment type.</param>
    public void SetAlignment(TextAlignment align)
    {
        if (currentParFormat.Alignment == align) return;
        
        var keyword = "";

        switch (align)
        { 
            case TextAlignment.Left:
                keyword = "ql";
                break;
            case TextAlignment.Right:
                keyword = "qr";
                break;
            case TextAlignment.Centered:
                keyword = "qc";
                break;
            case TextAlignment.Justified:
                keyword = "qj";
                break;
        }

        currentParFormat.Alignment = align;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, keyword, false, 0));
    }

    /// <summary>
    /// Sets the left indentation of the paragraph.
    /// </summary>
    /// <param name="val">Left indentation in centimeters.</param>
    public void SetLeftIndentation(float val)
    {
        if (Math.Abs(currentParFormat.LeftIndentation - val) < float.Epsilon) return;
        
        currentParFormat.LeftIndentation = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "li", true, calcTwips(val)));
    }

    /// <summary>
    /// Sets the right indentation of the paragraph.
    /// </summary>
    /// <param name="val">Right indentation in centimeters.</param>
    public void SetRightIndentation(float val)
    {
        if (Math.Abs(currentParFormat.RightIndentation - val) < float.Epsilon) return;
        
        currentParFormat.RightIndentation = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ri", true, calcTwips(val)));
    }

    /// <summary>
    /// Sets the bold font indicator.
    /// </summary>
    /// <param name="val">Bold font value</param>
    public void SetFormatBold(bool val)
    {
        if (currentFormat == null || currentFormat.Bold == val) return;
        currentFormat.Bold = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "b", !val, 0));
    }

    /// <summary>
    /// Set the italic/cursive font value.
    /// </summary>
    /// <param name="val">Cursive font value</param>
    public void SetFormatItalic(bool val)
    {
        if (currentFormat == null || currentFormat.Italic == val) return;
        currentFormat.Italic = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "i", !val, 0));
    }

    /// <summary>
    /// Set the underline font value
    /// </summary>
    /// <param name="val">Underline font value.</param>
    public void SetFormatUnderline(bool val)
    {
        if (currentFormat == null || currentFormat?.Underline == val) return;
        
        currentFormat!.Underline = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ul", !val, 0));
    }

    /// <summary>
    /// Set the font color value
    /// </summary>
    /// <param name="val">Color of the font.</param>
    public void SetFormatColor(SKColor val)
    {
        if (currentFormat == null || currentFormat?.Color == val) return;
        
        var indColor = colorTable.IndexOf(val);

        if (indColor == -1)
        {
            colorTable.AddColor(val);
            indColor = colorTable.IndexOf(val);
        }

        currentFormat!.Color = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "cf", true, indColor));
    }

    /// <summary>
    /// Set the font size
    /// </summary>
    /// <param name="val">Size of the font.</param>
    public void SetFormatSize(int val)
    {
        if (currentFormat == null || currentFormat?.Size == val) return;
        
        currentFormat!.Size = val;

        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fs", true, val * 2));
    }

    /// <summary>
    /// Sets the current font.
    /// </summary>
    /// <param name="val">Name of the font.</param>
    public void SetFormatFont(string val)
    {
        if (currentFormat == null || currentFormat.Font == val) return;
        var indFont = fontTable.IndexOf(val);

        if (indFont == -1)
        {
            fontTable.AddFont(val);
            indFont = fontTable.IndexOf(val);
        }

        currentFormat.Font = val;
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, indFont));
    }

    #endregion

    #region Properties

    /// <summary>
    /// Gets the plain text contained in the RTF document
    /// </summary>
    public string Text => GetTree().Text;

    /// <summary>
    /// Gets the RTF code of the RTF document
    /// </summary>
    public string Rtf => GetTree().Rtf;

    /// <summary>
    /// Gets the RTF tree of the current document
    /// </summary>
    public RtfTree Tree => GetTree();

    #endregion

    #region Private Methods

    /// <summary>
    /// Gets the RTF tree equivalent to the current document.
    /// <returns>RTF tree equivalent to the document in the current state.</returns>
    /// </summary>
    private RtfTree GetTree()
    {
        var tree = new RtfTree();
        tree.RootNode.AppendChild(mainGroup.CloneNode());

        InsertFontTable(tree);
        InsertColorTable(tree);
        InsertGenerator(tree);
        InsertDocSettings(tree);

        tree.MainGroup?.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));

        return tree;
    }
    
    /// <summary>
    /// Inserts the RTF code from the font table into the document.
    /// </summary>
    private void InsertFontTable(RtfTree tree)
    {
        if (tree.MainGroup == null) return;
        var ftGroup = new RtfTreeNode(RtfNodeType.Group);
        
        ftGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fonttbl", false, 0));

        for(var i=0; i<fontTable.Count; i++)
        {
            var ftFont = new RtfTreeNode(RtfNodeType.Group);
            ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, i));
            ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fnil", false, 0));
            ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Text, fontTable[i] + ";", false, 0));

            ftGroup.AppendChild(ftFont);
        }

        tree.MainGroup.InsertChild(5, ftGroup);
    }

    /// <summary>
    /// Insert the RTF code of the color table into the document.
    /// </summary>
    private void InsertColorTable(RtfTree tree)
    {
        if (tree.MainGroup == null) return;
        var ctGroup = new RtfTreeNode(RtfNodeType.Group);

        ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "colortbl", false, 0));

        for (var i = 0; i < colorTable.Count; i++)
        {
            ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "red", true, colorTable[i].Red));
            ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "green", true, colorTable[i].Green));
            ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "blue", true, colorTable[i].Blue));
            ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, ";", false, 0));
        }

        tree.MainGroup.InsertChild(6, ctGroup);
    }

    /// <summary>
    /// Insert the RTF code from the document generating application.
    /// </summary>
    private void InsertGenerator(RtfTree tree)
    {
        if (tree.MainGroup == null) return;
        var genGroup = new RtfTreeNode(RtfNodeType.Group);

        genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Control, "*", false, 0));
        genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "generator", false, 0));
        genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, "NRtfTree Library 0.3.0;", false, 0));

        tree.MainGroup.InsertChild(7, genGroup);
    }

    /// <summary>
    /// Inserts all the text and control nodes needed to represent a given text.
    /// </summary>
    /// <param name="text">Text to insert.</param>
    private void InsertText(string text)
    {
        var i = 0;
        var code = 0;

        while(i < text.Length)
        {
            code = char.ConvertToUtf32(text, i);

            if(code >= 32 && code < 128)
            {
                var s = new StringBuilder("");

                while (i < text.Length && code >= 32 && code < 128)
                {
                    s.Append(text[i]);

                    i++;

                    if(i < text.Length)
                        code = char.ConvertToUtf32(text, i);
                }

                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, s.ToString(), false, 0));
            }
            else
            {
                if (text[i] == '\t')
                {
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "tab", false, 0));
                }
                else if (text[i] == '\n')
                {
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
                }
                else
                {
                    if (code <= 255)
                    {
                        var bytes = encoding.GetBytes(new char[] { text[i] });

                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Control, "'", true, bytes[0]));
                    }
                    else
                    {
                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "u", true, code));
                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, "?", false, 0));
                    }
                }

                i++;
            }
        }
    }

    /// <summary>
    /// Update the font table with a new font if necessary.
    /// </summary>
    /// <param name="format"></param>
    private void UpdateFontTable(RtfCharFormat format)
    {
        if (fontTable.IndexOf(format.Font) == -1)
        {
            fontTable.AddFont(format.Font);
        }
    }

    /// <summary>
    /// Update the color table with a new color if necessary.
    /// </summary>
    /// <param name="format"></param>
    private void UpdateColorTable(RtfCharFormat format)
    {
        if (colorTable.IndexOf(format.Color) == -1)
        {
            colorTable.AddColor(format.Color);
        }
    }

    /// <summary>
    /// Initializes the RTF tree with all the keys from the document header.
    /// </summary>
    private void InitializeTree()
    {
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "rtf", true, 1));
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ansi", false, 0));
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ansicpg", true, encoding.CodePage));
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "deff", true, 0));
        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "deflang", true, CultureInfo.CurrentCulture.LCID));

        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
    }

    /// <summary>
    /// Inserts the document formatting properties
    /// </summary>
    private void InsertDocSettings(RtfTree tree)
    {
        if (tree.MainGroup == null || tree.MainGroup.ChildNodes == null) return;
        var indInicioTexto = tree.MainGroup.ChildNodes.IndexOf("pard");

        //Generic Properties
        
        tree.MainGroup.InsertChild(indInicioTexto, new RtfTreeNode(RtfNodeType.Keyword, "viewkind", true, 4));
        tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "uc", true, 1));

        //RtfDocumentFormat Properties

        tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margl", true, calcTwips(docFormat.MarginL)));
        tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margr", true, calcTwips(docFormat.MarginR)));
        tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margt", true, calcTwips(docFormat.MarginT)));
        tree.MainGroup.InsertChild(indInicioTexto  , new RtfTreeNode(RtfNodeType.Keyword, "margb", true, calcTwips(docFormat.MarginB)));
    }

    /// <summary>
    /// Convert between centimeters and twips.
    /// </summary>
    /// <param name="centimeters">Value in centimeters.</param>
    /// <returns>Value in twips.</returns>
    private int calcTwips(float centimeters)
    {
        // DPI dependant?
        //1 inches = 2.54 centimeters
        //1 inches = 1440 twips
        //20 twip = 1 pixel

        //X centimeters --> (X*1440)/2.54 twips

        return (int)((centimeters * 1440F) / 2.54F);
    }

    #endregion
}

