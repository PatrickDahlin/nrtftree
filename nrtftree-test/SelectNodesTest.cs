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
 * Class:		SelectNodesTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree.Test;

public class SelectNodesTest
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
    public void SelectChildNodesByType()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectChildNodes(RtfNodeType.Keyword);  //48 nodes
        var lista2 = tree.MainGroup.SelectChildNodes(RtfNodeType.Control);  //3 nodes
        var lista3 = tree.MainGroup.SelectChildNodes(RtfNodeType.Group);    //3 nodes

        Assert.That(lista1.Count, Is.EqualTo(49));
        Assert.That(lista2.Count, Is.EqualTo(3));
        Assert.That(lista3.Count, Is.EqualTo(3));

        Assert.That(lista1[5], Is.SameAs(tree.MainGroup[8]));  //viewkind
        Assert.That(lista1[22].NodeKey, Is.EqualTo("lang"));   //lang3082  

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[45])); //'233
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[47])); //'241
        Assert.That(lista2[1].Parameter, Is.EqualTo(241));     //'241

        Assert.That(lista3[0], Is.SameAs(tree.MainGroup[5]));
        Assert.That(lista3[0].FirstChild.NodeKey, Is.EqualTo("fonttbl"));
        Assert.That(lista3[1], Is.SameAs(tree.MainGroup[6]));
        Assert.That(lista3[1].FirstChild.NodeKey, Is.EqualTo("colortbl"));
        Assert.That(lista3[2], Is.SameAs(tree.MainGroup[7]));
        Assert.That(lista3[2].ChildNodes[0].NodeKey, Is.EqualTo("*"));
        Assert.That(lista3[2].ChildNodes[1].NodeKey, Is.EqualTo("generator"));
    }

    [Test]
    public void SelectNodesByType()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectNodes(RtfNodeType.Keyword);  //68 nodes
        var lista2 = tree.MainGroup.SelectNodes(RtfNodeType.Control);  //4 nodes
        var lista3 = tree.MainGroup.SelectNodes(RtfNodeType.Group);    //6 nodes

        Assert.That(lista1.Count, Is.EqualTo(69));
        Assert.That(lista2.Count, Is.EqualTo(4));
        Assert.That(lista3.Count, Is.EqualTo(6));

        Assert.That(lista1[5], Is.SameAs(tree.MainGroup[5].FirstChild));     //fonttbl
        Assert.That(lista1[22], Is.SameAs(tree.MainGroup[6].ChildNodes[7])); //green0
        Assert.That(lista1[22].NodeKey, Is.EqualTo("green"));                //green0  

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[7].FirstChild)); //* generator
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[45])); //'233
        Assert.That(lista2[2], Is.SameAs(tree.MainGroup[47])); //'241
        Assert.That(lista2[2].Parameter, Is.EqualTo(241));     //'241

        Assert.That(lista3[0], Is.SameAs(tree.MainGroup[5]));
        Assert.That(lista3[0].FirstChild.NodeKey, Is.EqualTo("fonttbl"));
        Assert.That(lista3[3], Is.SameAs(tree.MainGroup[5].ChildNodes[3]));
        Assert.That(lista3[3].FirstChild.NodeKey, Is.EqualTo("f"));
        Assert.That(lista3[3].FirstChild.Parameter, Is.EqualTo(2));
        Assert.That(lista3[5], Is.SameAs(tree.MainGroup[7]));
        Assert.That(lista3[5].ChildNodes[0].NodeKey, Is.EqualTo("*"));
        Assert.That(lista3[5].ChildNodes[1].NodeKey, Is.EqualTo("generator"));
    }

    [Test]
    public void SelectSingleNodeByType()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleNode(RtfNodeType.Keyword); //rtf1
        var node2 = tree.MainGroup.SelectSingleNode(RtfNodeType.Control); //* generator
        var node3 = tree.MainGroup.SelectSingleNode(RtfNodeType.Group);   //fonttbl

        Assert.That(node1, Is.SameAs(tree.MainGroup[0]));
        Assert.That(node1.NodeKey, Is.EqualTo("rtf"));
        Assert.That(node2, Is.SameAs(tree.MainGroup[7].ChildNodes[0]));
        Assert.That(node2.NodeKey, Is.EqualTo("*"));
        Assert.That(node2.NextSibling.NodeKey, Is.EqualTo("generator"));
        Assert.That(node3, Is.SameAs(tree.MainGroup[5]));
        Assert.That(node3.FirstChild.NodeKey, Is.EqualTo("fonttbl"));
    }

    [Test]
    public void SelectSingleChildNodeByType()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleChildNode(RtfNodeType.Keyword); //rtf1
        var node2 = tree.MainGroup.SelectSingleChildNode(RtfNodeType.Control); //'233
        var node3 = tree.MainGroup.SelectSingleChildNode(RtfNodeType.Group);   //fonttbl

        Assert.That(node1, Is.SameAs(tree.MainGroup[0]));
        Assert.That(node1.NodeKey, Is.EqualTo("rtf"));
        Assert.That(node2, Is.SameAs(tree.MainGroup[45]));
        Assert.That(node2.NodeKey, Is.EqualTo("'"));
        Assert.That(node2.Parameter, Is.EqualTo(233));
        Assert.That(node3, Is.SameAs(tree.MainGroup[5]));
        Assert.That(node3.FirstChild.NodeKey, Is.EqualTo("fonttbl"));
    }

    [Test]
    public void SelectChildNodesByKeyword()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectChildNodes("fs");  //5 nodes
        var lista2 = tree.MainGroup.SelectChildNodes("f");   //3 nodes

        Assert.That(lista1.Count, Is.EqualTo(5));
        Assert.That(lista2.Count, Is.EqualTo(3));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[17]));
        Assert.That(lista1[1], Is.SameAs(tree.MainGroup[22]));
        Assert.That(lista1[2], Is.SameAs(tree.MainGroup[25]));
        Assert.That(lista1[3], Is.SameAs(tree.MainGroup[43]));
        Assert.That(lista1[4], Is.SameAs(tree.MainGroup[77]));

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[16]));
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[56]));
        Assert.That(lista2[2], Is.SameAs(tree.MainGroup[76]));
    }

    [Test]
    public void SelectNodesByKeyword()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectNodes("fs");  //5 nodes
        var lista2 = tree.MainGroup.SelectNodes("f");   //6 nodes

        Assert.That(lista1.Count, Is.EqualTo(5));
        Assert.That(lista2.Count, Is.EqualTo(6));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[17]));
        Assert.That(lista1[1], Is.SameAs(tree.MainGroup[22]));
        Assert.That(lista1[2], Is.SameAs(tree.MainGroup[25]));
        Assert.That(lista1[3], Is.SameAs(tree.MainGroup[43]));
        Assert.That(lista1[4], Is.SameAs(tree.MainGroup[77]));

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[5].ChildNodes[1].FirstChild));
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[5].ChildNodes[2].FirstChild));
        Assert.That(lista2[2], Is.SameAs(tree.MainGroup[5].ChildNodes[3].FirstChild));
        Assert.That(lista2[3], Is.SameAs(tree.MainGroup[16]));
        Assert.That(lista2[4], Is.SameAs(tree.MainGroup[56]));
        Assert.That(lista2[5], Is.SameAs(tree.MainGroup[76]));
    }

    [Test]
    public void SelectSingleNodeByKeyword()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleNode("fs"); 
        var node2 = tree.MainGroup.SelectSingleNode("f");  

        Assert.That(node1, Is.SameAs(tree.MainGroup[17]));
        Assert.That(node2, Is.SameAs(tree.MainGroup[5].ChildNodes[1].FirstChild));
    }

    [Test]
    public void SelectSingleChildNodeByKeyword()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleChildNode("fs");
        var node2 = tree.MainGroup.SelectSingleChildNode("f");

        Assert.That(node1, Is.SameAs(tree.MainGroup[17]));
        Assert.That(node2, Is.SameAs(tree.MainGroup[16]));
    }

    [Test]
    public void SelectChildNodesByKeywordAndParam()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectChildNodes("fs", 24);  //2 nodes
        var lista2 = tree.MainGroup.SelectChildNodes("f", 1);    //1 nodes

        Assert.That(lista1.Count, Is.EqualTo(2));
        Assert.That(lista2.Count, Is.EqualTo(1));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[22]));
        Assert.That(lista1[1], Is.SameAs(tree.MainGroup[43]));

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[56]));
    }

    [Test]
    public void SelectNodesByKeywordAndParam()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectNodes("fs", 24);  //2 nodes
        var lista2 = tree.MainGroup.SelectNodes("f", 1);    //2 nodes

        Assert.That(lista1.Count, Is.EqualTo(2));
        Assert.That(lista2.Count, Is.EqualTo(2));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[22]));
        Assert.That(lista1[1], Is.SameAs(tree.MainGroup[43]));

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[5].ChildNodes[2].FirstChild));
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[56]));
    }

    [Test]
    public void SelectSingleNodeByKeywordAndParam()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleNode("fs", 24);
        var node2 = tree.MainGroup.SelectSingleNode("f", 1);

        Assert.That(node1, Is.SameAs(tree.MainGroup[22]));
        Assert.That(node2, Is.SameAs(tree.MainGroup[5].ChildNodes[2].FirstChild));
    }

    [Test]
    public void SelectSingleChildNodeByKeywordAndParam()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleChildNode("fs", 24);
        var node2 = tree.MainGroup.SelectSingleChildNode("f", 1);

        Assert.That(node1, Is.SameAs(tree.MainGroup[22]));
        Assert.That(node2, Is.SameAs(tree.MainGroup[56]));
    }

    [Test]
    public void SelectChildGroups()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectChildGroups("colortbl");  //1 node
        var lista2 = tree.MainGroup.SelectChildGroups("f");         //0 nodes

        Assert.That(lista1.Count, Is.EqualTo(1));
        Assert.That(lista2.Count, Is.EqualTo(0));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[6]));
    }

    [Test]
    public void SelectGroups()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var lista1 = tree.MainGroup.SelectGroups("colortbl");  //1 node
        var lista2 = tree.MainGroup.SelectGroups("f");         //3 nodes

        Assert.That(lista1.Count, Is.EqualTo(1));
        Assert.That(lista2.Count, Is.EqualTo(3));

        Assert.That(lista1[0], Is.SameAs(tree.MainGroup[6]));

        Assert.That(lista2[0], Is.SameAs(tree.MainGroup[5].ChildNodes[1]));
        Assert.That(lista2[1], Is.SameAs(tree.MainGroup[5].ChildNodes[2]));
        Assert.That(lista2[2], Is.SameAs(tree.MainGroup[5].ChildNodes[3]));
    }

    [Test]
    public void SelectSingleGroup()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.SelectSingleGroup("f");
        var node2 = tree.MainGroup[5].SelectSingleChildGroup("f");

        Assert.That(node1, Is.SameAs(tree.MainGroup[5].ChildNodes[1]));
        Assert.That(node2, Is.SameAs(tree.MainGroup[5].ChildNodes[1]));
    }

    [Test]
    public void SelectSpecialGroups()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var list1 = tree.MainGroup.SelectChildGroups("generator");
        var list2 = tree.MainGroup.SelectChildGroups("generator", false);
        var list3 = tree.MainGroup.SelectChildGroups("generator", true);

        var list4 = tree.MainGroup.SelectGroups("generator");
        var list5 = tree.MainGroup.SelectGroups("generator", false);
        var list6 = tree.MainGroup.SelectGroups("generator", true);

        var node1 = tree.MainGroup.SelectSingleChildGroup("generator");
        var node2 = tree.MainGroup.SelectSingleChildGroup("generator", false);
        var node3 = tree.MainGroup.SelectSingleChildGroup("generator", true);

        var node4 = tree.MainGroup.SelectSingleGroup("generator");
        var node5 = tree.MainGroup.SelectSingleGroup("generator", false);
        var node6 = tree.MainGroup.SelectSingleGroup("generator", true);

        Assert.That(list1.Count, Is.EqualTo(0));
        Assert.That(list2.Count, Is.EqualTo(0));
        Assert.That(list3.Count, Is.EqualTo(1));

        Assert.That(list4.Count, Is.EqualTo(0));
        Assert.That(list5.Count, Is.EqualTo(0));
        Assert.That(list6.Count, Is.EqualTo(1));

        Assert.That(node1, Is.Null);
        Assert.That(node2, Is.Null);
        Assert.That(node3, Is.Not.Null);

        Assert.That(node4, Is.Null);
        Assert.That(node5, Is.Null);
        Assert.That(node6, Is.Not.Null);
    }

    [Test]
    public void SelectSiblings()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var node1 = tree.MainGroup.ChildNodes[4];               //deflang3082
        var node2 = tree.MainGroup.ChildNodes[6].ChildNodes[2]; //colortbl/red

        var n1 = node1.SelectSibling(RtfNodeType.Group);
        var n2 = node1.SelectSibling("viewkind");
        var n3 = node1.SelectSibling("fs", 28);

        var n4 = node2.SelectSibling(RtfNodeType.Keyword);
        var n5 = node2.SelectSibling("blue");
        var n6 = node2.SelectSibling("red", 255);

        Assert.That(n1, Is.SameAs(tree.MainGroup[5]));
        Assert.That(n2, Is.SameAs(tree.MainGroup[8]));
        Assert.That(n3, Is.SameAs(tree.MainGroup[17]));

        Assert.That(n4, Is.SameAs(tree.MainGroup[6].ChildNodes[3]));
        Assert.That(n5, Is.SameAs(tree.MainGroup[6].ChildNodes[4]));
        Assert.That(n6, Is.SameAs(tree.MainGroup[6].ChildNodes[6]));
    }

    [Test]
    public void FindText()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var list1 = tree.MainGroup.FindText("Italic");

        Assert.That(list1.Count, Is.EqualTo(2));

        Assert.That(list1[0], Is.SameAs(tree.MainGroup[18]));
        Assert.That(list1[0].NodeKey, Is.EqualTo("Bold Italic Underline Size 14"));

        Assert.That(list1[1], Is.SameAs(tree.MainGroup[73]));
        Assert.That(list1[1].NodeKey, Is.EqualTo("Italic2"));
    }

    [Test]
    public void ReplaceText()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        tree.MainGroup.ReplaceText("Italic", "REPLACED");

        var sr = new StreamReader(@"..\..\..\testdocs\rtf2.txt");
        var rtf2 = sr.ReadToEnd();
        sr.Close();

        Assert.That(tree.Rtf, Is.EqualTo(rtf2));
    }
}

