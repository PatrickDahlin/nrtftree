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
 * Class:		ObjectNodeTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;

namespace Net.Sgoliver.NRtfTree.Test;

public class ObjectNodeTest
{

    [SetUp]
    public void Setup()
    {
#if NETCORE || NET
        // Add a reference to the NuGet package System.Text.Encoding.CodePages for .Net core only
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif   
    }

    [Test]
    public void LoadObjectNode()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var node = tree.MainGroup.SelectSingleNode("object").ParentNode;

        var objNode = new ObjectNode(node);

        Assert.That(objNode.ObjectType, Is.EqualTo("objemb"));
        Assert.That(objNode.ObjectClass, Is.EqualTo("Excel.Sheet.8"));
    }

    [Test]
    public void ObjectHexData()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var node = tree.MainGroup.SelectSingleNode("object").ParentNode;

        var objNode = new ObjectNode(node);

        var sr = new StreamReader(@"..\..\..\testdocs\objhexdata.txt");
        var hexdata = sr.ReadToEnd();
        sr.Close();

        Assert.That(objNode.HexData, Is.EqualTo(hexdata));
    }

    [Test]
    public void ObjectBinData()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var node = tree.MainGroup.SelectSingleNode("object").ParentNode;

        var objNode = new ObjectNode(node);

        var bw = new BinaryWriter(new FileStream(@"..\..\..\testdocs\objbindata-result.dat", FileMode.Create));
        foreach (var b in objNode.GetByteData())
            bw.Write(b);
        bw.Close();

        var fs1 = new FileStream(@"..\..\..\testdocs\objbindata-result.dat", FileMode.Open);
        var fs2 = new FileStream(@"..\..\..\testdocs\objbindata.dat", FileMode.Open);

        Assert.That(fs1, Is.EqualTo(fs2));
    }

    [Test]
    public void ResultNode()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var node = tree.MainGroup.SelectSingleNode("object").ParentNode;

        var objNode = new ObjectNode(node);

        var resNode = objNode.ResultNode;

        Assert.That(resNode, Is.SameAs(tree.MainGroup.SelectSingleGroup("object").SelectSingleChildGroup("result")));

        var pictNode = resNode.SelectSingleNode("pict").ParentNode;
        var imgNode = new ImageNode(pictNode);

        Assert.That(imgNode.Height, Is.EqualTo(2247));
        Assert.That(imgNode.Width, Is.EqualTo(9320));

        Assert.That(imgNode.DesiredHeight, Is.EqualTo(1274));
        Assert.That(imgNode.DesiredWidth, Is.EqualTo(5284));

        Assert.That(imgNode.ScaleX, Is.EqualTo(100));
        Assert.That(imgNode.ScaleY, Is.EqualTo(100));

        Assert.That(imgNode.ImageFormat, Is.EqualTo(ImageFormat.Emf));
    }
}

