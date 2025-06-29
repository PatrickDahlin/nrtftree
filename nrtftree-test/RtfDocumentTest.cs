﻿/********************************************************************************
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
 * Class:		RtfDocumentTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using Net.Sgoliver.NRtfTree.Util;
using System.Globalization;
using System.Text;
using SkiaSharp;

namespace Net.Sgoliver.NRtfTree.Test;

public class RtfDocumentTest
{

    [SetUp]
    public void InitTest()
    {
#if NETCORE || NET
        // Add a reference to the NuGet package System.Text.Encoding.CodePages for .Net core only
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
    }

    [Test]
    public void CreateSimpleDocument()
    {
        var doc = new RtfDocument(Encoding.GetEncoding(1252)); // Encoding.Default has changed to UTF8 in newer .NETs

        var charFormat = new RtfCharFormat();
        charFormat.Color = SKColors.DarkBlue;
        charFormat.Underline = true;
        charFormat.Bold = true;
        doc.UpdateCharFormat(charFormat);

        var parFormat = new RtfParFormat();
        parFormat.Alignment = TextAlignment.Justified;
        doc.UpdateParFormat(parFormat);

        doc.AddText("First Paragraph");
        doc.AddNewParagraph(2);

        doc.SetFormatBold(false);
        doc.SetFormatUnderline(false);
        doc.SetFormatColor(SKColors.Red);

        doc.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer quis eros at tortor pharetra laoreet. Donec tortor diam, imperdiet ut porta quis, congue eu justo.");
        doc.AddText("Quisque viverra tellus id mauris tincidunt luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.");
        doc.AddText("Donec ac leo justo, vitae rutrum elit. Nulla tellus elit, imperdiet luctus porta vel, consectetur quis turpis. Nam purus odio, dictum vitae sollicitudin nec, tempor eget mi.");
        doc.AddText("Etiam vitae porttitor enim. Aenean molestie facilisis magna, quis tincidunt leo placerat in. Maecenas malesuada eleifend nunc vitae cursus.");
        doc.AddNewParagraph(2);
        
        doc.Save(@"..\..\..\testdocs\rtfdocument1.rtf");

        var text1 = doc.Text;
        var rtfcode1 = doc.Rtf;

        doc.AddText("Second Paragraph", charFormat);
        doc.AddNewParagraph(2);

        charFormat.Font = "Courier New";
        charFormat.Color = SKColors.Green;
        charFormat.Bold = false;
        charFormat.Underline = false;
        doc.UpdateCharFormat(charFormat);

        doc.SetAlignment(TextAlignment.Left);
        doc.SetLeftIndentation(2);
        doc.SetRightIndentation(2);

        doc.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer quis eros at tortor pharetra laoreet. Donec tortor diam, imperdiet ut porta quis, congue eu justo.");
        doc.AddText("Quisque viverra tellus id mauris tincidunt luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.");
        doc.AddText("Donec ac leo justo, vitae rutrum elit. Nulla tellus elit, imperdiet luctus porta vel, consectetur quis turpis. Nam purus odio, dictum vitae sollicitudin nec, tempor eget mi.");
        doc.AddText("Etiam vitae porttitor enim. Aenean molestie facilisis magna, quis tincidunt leo placerat in. Maecenas malesuada eleifend nunc vitae cursus.");
        doc.AddNewParagraph(2);

        doc.UpdateCharFormat(charFormat);
        doc.SetFormatUnderline(false);
        doc.SetFormatItalic(true);
        doc.SetFormatColor(SKColors.DarkBlue);

        doc.SetLeftIndentation(0);

        doc.AddText("Test Doc. Петяв ñáéíó\n");
        doc.AddNewLine(1);
        doc.AddText("\tStop.");

        var text2 = doc.Text;
        var rtfcode2 = doc.Rtf;

        doc.Save(@"..\..\..\testdocs\rtfdocument2.rtf");

        var sr = new StreamReader(@"..\..\..\testdocs\rtfdocument1.rtf");
        var rtf1 = sr.ReadToEnd();
        sr.Close();

        sr = null;
        sr = new StreamReader(@"..\..\..\testdocs\rtfdocument2.rtf");
        var rtf2 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf4.txt");
        var rtf4 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf6.txt");
        var rtf6 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\doctext1.txt");
        var doctext1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\doctext2.txt");
        var doctext2 = sr.ReadToEnd() + " Петяв ñáéíó\r\n\r\n\tStop.\r\n";
        sr.Close();

        //Se adapta el lenguaje al del PC donde se ejecutan los tests
        var deflangInd = rtf4.IndexOf("\\deflang3082");
        rtf4 = rtf4.Substring(0, deflangInd) + "\\deflang" + CultureInfo.CurrentCulture.LCID + rtf4.Substring(deflangInd + 8 + CultureInfo.CurrentCulture.LCID.ToString().Length);

        //Se adapta el lenguaje al del PC donde se ejecutan los tests
        var deflangInd2 = rtf6.IndexOf("\\deflang3082");
        rtf6 = rtf6.Substring(0, deflangInd2) + "\\deflang" + CultureInfo.CurrentCulture.LCID + rtf6.Substring(deflangInd2 + 8 + CultureInfo.CurrentCulture.LCID.ToString().Length);

        Assert.That(rtf1, Is.EqualTo(rtf6));
        Assert.That(rtf2, Is.EqualTo(rtf4));

        Assert.That(text1, Is.EqualTo(doctext1));
        Assert.That(text2, Is.EqualTo(doctext2));
    }
}
