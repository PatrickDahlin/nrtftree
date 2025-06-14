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
 * Class:		Demo
 * Description:	Programa principal de demostración de la librería.
 ********************************************************************************/

using System.Text;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Demo;
using Net.Sgoliver.NRtfTree.Util;

public class Demo
{

    static void Main(string[] args)
    {
#if NETCORE || NET
        // Add a reference to the NuGet package System.Text.Encoding.CodePages for .Net core only
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif  
        var option = "";

        var language = SelectLanguage();

        while (!option.Equals("0"))
        {
            if (language == 1)
                PrintEnglishMenu();
            else if (language == 2)
                PrintSpanishMenu();

            option = Console.In.ReadLine();

            Console.WriteLine("");

            switch (option)
            {
                case "1":
                    ExtractDocumentProperties();
                    break;
                case "2":
                    GenerateRtfTree();
                    break;
                case "3":
                    ExtractPlainText();
                    break;
                case "4":
                    ExtractDocumentOutline();
                    break;
                case "5":
                    ExtractHyperlinks();
                    break;
                case "6":
                    ExtractImages();
                    break;
                case "7":
                    ExtractObjects();
                    break;
                case "8":
                    TagFormat();
                    break;
                case "9":
                    MergeDocuments();
                    break;
                case "10":
                    ConvertToHtml();
                    break;
            }
        }
    }
    
    private static void PrintEnglishMenu()
    {
        Console.WriteLine("");
        Console.WriteLine("Select Demo:");
        Console.WriteLine("");
        Console.WriteLine("   1. Extract document properties.");
        Console.WriteLine("   2. Generate RTF tree.");
        Console.WriteLine("   3. Extract plain text.");
        Console.WriteLine("   4. Extract document outline.");
        Console.WriteLine("   5. Extract hyperlinks.");
        Console.WriteLine("   6. Extract images.");
        Console.WriteLine("   7. Extract embedded objects.");
        Console.WriteLine("   8. Tag format (RtfReader Example).");
        Console.WriteLine("   9. Merge documents.");
        Console.WriteLine("  10. Translate to HTML.");
        Console.WriteLine("");
        Console.WriteLine("   0. Quit.");
        Console.WriteLine("");
        Console.Write("Option> ");
    }
    
    private static void PrintSpanishMenu()
    {
        Console.WriteLine("");
        Console.WriteLine("Selecciona ejemplo:");
        Console.WriteLine("");
        Console.WriteLine("   1. Extraer propiedades del documento.");
        Console.WriteLine("   2. Generar árbol RTF.");
        Console.WriteLine("   3. Extraer texto plano.");
        Console.WriteLine("   4. Extraer esquema del documento.");
        Console.WriteLine("   5. Extraer enlaces.");
        Console.WriteLine("   6. Extraer imágenes.");
        Console.WriteLine("   7. Extraer objetos incrustados.");
        Console.WriteLine("   8. Etiquetar formato (Ejemplo de RtfReader).");
        Console.WriteLine("   9. Combinar documentos.");
        Console.WriteLine("  10. Convertir a HTML.");
        Console.WriteLine("");
        Console.WriteLine("   0. Salir.");
        Console.WriteLine("");
        Console.Write("Opción> ");
    }
    
    private static int SelectLanguage()
    {
        var language = 0;

        PrintLanguageMenu();

        var option = Console.In.ReadLine();

        Console.WriteLine("");

        switch (option)
        {
            case "1":
                language = 1;
                break;
            case "2":
                language = 2;
                break;
        }

        return language;
    }
    
    private static void PrintLanguageMenu()
    {
        Console.WriteLine("");
        Console.WriteLine("Select Language / Selecciona Idioma:");
        Console.WriteLine("");
        Console.WriteLine("   1. English.");
        Console.WriteLine("   2. Español.");
        Console.WriteLine("");
        Console.Write("> ");
    }
    
    public static void ExtractDocumentProperties()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var info = tree.GetInfoGroup();

        Console.WriteLine("Extracting document properties:");

        Console.WriteLine("Title: {0}", info.Title);
        Console.WriteLine("Author: {0}", info.Author);
        Console.WriteLine("Company: {0}", info.Company);
        Console.WriteLine("Comments: {0}", info.DocComment);
        Console.WriteLine("Created: {0}", info.CreationTime);
        Console.WriteLine("Revised: {0}", info.RevisionTime);

        Console.WriteLine("");
    }
    
    public static void GenerateRtfTree()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var sw = new StreamWriter(@"..\..\..\testdocs\rtftree.txt");

        Console.WriteLine("Generating RTF tree...");

        sw.Write(tree.ToStringEx());
        sw.Flush();
        sw.Close();

        Console.WriteLine("File 'rtftree.txt' created.");

        Console.WriteLine("");
    }
    
    public static void ExtractPlainText()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var sw = new StreamWriter(@"..\..\..\testdocs\rtftext.txt");

        Console.WriteLine("Extracting text...");

        sw.Write(tree.Text);
        sw.Flush();
        sw.Close();

        Console.WriteLine("File 'rtftext.txt' created.");

        Console.WriteLine("");
    }
    
    public static void ExtractDocumentOutline()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var sst = tree.GetStyleSheetTable();

        var heading1 = sst.IndexOf("heading 1");
        var heading2 = sst.IndexOf("heading 2");
        var heading3 = sst.IndexOf("heading 3");

        tree.MainGroup.RemoveChild(tree.MainGroup.SelectChildGroups("stylesheet")[0]);

        var headingKeywords = tree.MainGroup.SelectNodes("s");

        for (var i = 0; i < headingKeywords.Count; i++)
        {
            var hk = headingKeywords[i];

            var text = new StringBuilder("");

            if (hk.Parameter == heading1 ||
                hk.Parameter == heading2 ||
                hk.Parameter == heading3)
            {
                var sibling = hk.NextSibling;

                while (sibling != null && !sibling.NodeKey.Equals("pard"))
                {
                    if (sibling.NodeType == RtfNodeType.Text)
                        text.Append(sibling.NodeKey);
                    else if (sibling.NodeType == RtfNodeType.Group)
                        text.Append(ExtractGroupText(sibling));

                    sibling = sibling.NextSibling;
                }

                if (hk.Parameter == heading1)
                    Console.WriteLine("H1: {0}", text);
                else if (hk.Parameter == heading2)
                    Console.WriteLine("    H2: {0}", text);
                else if (hk.Parameter == heading3)
                    Console.WriteLine("        H3: {0}", text);
            }
        }
    }
    
    private static string ExtractGroupText(RtfTreeNode group)
    {
        var sb = new StringBuilder("");

        foreach (RtfTreeNode node in group.ChildNodes)
        {
            if (node.NodeType == RtfNodeType.Text)
                sb.Append(node.NodeKey);
            else if (node.NodeType == RtfNodeType.Group)
                sb.Append(ExtractGroupText(node));
        }

        return sb.ToString();
    }
    
    private static void ExtractHyperlinks()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var fields = tree.MainGroup.SelectGroups("field");

        foreach (RtfTreeNode node in fields)
        {
            //Extract URL

            var fldInst = node.ChildNodes[1];

            var fldInstText = ExtractGroupText(fldInst);

            var url = fldInstText.Split([" "], StringSplitOptions.RemoveEmptyEntries)[1];

            //Extract Link Text

            var fldRslt = node.SelectSingleChildGroup("fldrslt");

            var linkText = ExtractGroupText(fldRslt);

            Console.WriteLine("[" + linkText + ", " + url + "]");
        }
    }
    
    private static void ExtractImages()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        var imageNodes = tree.RootNode.SelectNodes("pict");

        Console.WriteLine("Extracting images...");

        var i = 1;
        foreach (RtfTreeNode node in imageNodes)
        {
            var imageNode = new ImageNode(node.ParentNode);

            if (imageNode.ImageFormat == Net.Sgoliver.NRtfTree.Util.ImageFormat.Png)
            {
                imageNode.SaveImage(@"..\..\..\testdocs\image" + i + ".png");

                Console.WriteLine("File '" + "image" + i + ".png" + "' created.");

                i++;
            }
        }

        Console.WriteLine("");
    }
    
    private static void ExtractObjects()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc.rtf");

        //Busca el primer nodo de tipo objeto.
        var objects = tree.RootNode.SelectGroups("object");

        Console.WriteLine("Extracting objects...");

        var i = 1;
        foreach (RtfTreeNode? node in objects)
        {
            //Se crea un nodo RTF especializado en imágenes
            var objectNode = new ObjectNode(node);

            Console.WriteLine("Found new object:");
            Console.WriteLine("Object type: " + objectNode.ObjectType);
            Console.WriteLine("Object class: " + objectNode.ObjectClass);

            var data = objectNode.GetByteData();

            var binaryFile = new FileStream(@"..\..\..\testdocs\object" + i + ".xls", FileMode.Create, FileAccess.ReadWrite);
            var bw = new BinaryWriter(binaryFile);

            for (var j = 38; j < data.Length; j++)
            {
                bw.Write(data[j]);
            }
            bw.Flush();
            bw.Close();

            Console.WriteLine("File 'object" + i + ".xls' created.");

            i++;
        }

        Console.WriteLine("");
    }
    
    private static void TagFormat()
    {
        var res = "";

        var parser = new MyParser(res);

        var reader = new RtfReader(parser);

        reader.LoadRtfFile(@"..\..\..\testdocs\test-doc2.rtf");

        Console.WriteLine("Processing...");

        reader.Parse();

        var sw = new StreamWriter(@"..\..\..\testdocs\taggedfile.txt");
        sw.Write(parser.doc);
        sw.Flush();
        sw.Close();

        Console.WriteLine("File 'taggedfile.txt' created.");

        Console.WriteLine("");
    }
    
    private static void MergeDocuments()
    {
        var merger = new RtfMerger(@"..\..\..\testdocs\test-doc3.rtf");

        merger.AddPlaceHolder("[TagTextRTF1]", @"..\..\..\testdocs\merge1.rtf");
        merger.AddPlaceHolder("[TagTextRTF2]", @"..\..\..\testdocs\merge2.rtf");

        Console.WriteLine("Processing...");

        var tree = merger.Merge();
        tree.SaveRtf(@"..\..\..\testdocs\merge-result.rtf");

        Console.WriteLine("File 'merge-result.txt' created.");

        Console.WriteLine("");
    }
    
    private static void ConvertToHtml()
    {
        var tree = new RtfTree();
        tree.LoadRtfFile(@"..\..\..\testdocs\test-doc2.rtf");

        var rtfToHtml = new Rtf2Html();

        Console.WriteLine("Processing...");
        rtfToHtml.IncrustImages = false;
        var html = rtfToHtml.Convert(tree.Rtf);

        var sw = new StreamWriter(@"..\..\..\testdocs\test.html", false);
        sw.Write(html);
        sw.Flush();
        sw.Close();

        Console.WriteLine("File 'test.html' created.");

        Console.WriteLine("");
    }
    
}