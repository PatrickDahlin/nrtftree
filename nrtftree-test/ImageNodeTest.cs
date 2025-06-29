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
 * Class:		ImageNodeTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;

namespace Net.Sgoliver.NRtfTree.Test;

public class ImageNodeTest
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
    public void LoadImageNode()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

        var imgNode = new ImageNode(pictNode);

        Assert.That(imgNode.Height, Is.EqualTo(6615));
        Assert.That(imgNode.Width, Is.EqualTo(7938));

        Assert.That(imgNode.DesiredHeight, Is.EqualTo(3750));
        Assert.That(imgNode.DesiredWidth, Is.EqualTo(4500));

        Assert.That(imgNode.ScaleX, Is.EqualTo(100));
        Assert.That(imgNode.ScaleY, Is.EqualTo(100));

        Assert.That(imgNode.ImageFormat, Is.EqualTo(ImageFormat.Png));
    }

    [Test]
    public void ImageHexData()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

        var imgNode = new ImageNode(pictNode);

        StreamReader sr = null;

        sr = new StreamReader(@"..\..\..\testdocs\imghexdata.txt");
        var hexdata = sr.ReadToEnd();
        sr.Close();

        Assert.That(imgNode.HexData, Is.EqualTo(hexdata));
    }

    [Test]
    public void ImageBinData()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

        var imgNode = new ImageNode(pictNode);

        imgNode.SaveImage(@"..\..\..\testdocs\img-result.png", ImageFormat.Jpeg);

        Stream fs1 = new FileStream(@"..\..\..\testdocs\img-result.jpg", FileMode.Open);
        Stream fs2 = new FileStream(@"..\..\..\testdocs\image1.jpg", FileMode.Open);

        Assert.That(fs1, Is.EqualTo(fs2));
    }
}

