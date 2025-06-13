using System.Text;
using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree.Test;

public class LoadRtfTest
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
    public void LoadSimpleDocFromFile()
    {

        var tree = new RtfTree();
        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var sr = new StreamReader(@"..\..\..\testdocs\result1-1.txt");
        var strTree1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\result1-2.txt");
        var strTree2 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf1.txt");
        var rtf1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\text1.txt");
        var text1 = sr.ReadToEnd();
        sr.Close();

        Assert.Multiple(() =>
        {
            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.False);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        });
    }
    
    [Test]
    public void LoadImageDocFromFile()
    {
        var tree = new RtfTree();

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc3.rtf");

        var sr = new StreamReader(@"..\..\..\testdocs\rtf5.txt");
        var rtf5 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\text2.txt");
        var text2 = sr.ReadToEnd();
        sr.Close();

        Assert.Multiple(() =>
        {
            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.False);
            Assert.That(tree.Rtf, Is.EqualTo(rtf5));
            Assert.That(tree.Text, Is.EqualTo(text2));
        });
    }
    
    [Test]
    public void LoadSimpleDocMergeSpecialFromFile()
    {
        var tree = new RtfTree();

        tree.MergeSpecialCharacters = true;

        var res = tree.LoadRtfFile(@"..\..\..\testdocs\testdoc1.rtf");

        var sr = new StreamReader(@"..\..\..\testdocs\result1-3.txt");
        var strTree1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\result1-4.txt");
        var strTree2 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf1.txt");
        var rtf1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\text1.txt");
        var text1 = sr.ReadToEnd();
        sr.Close();

        Assert.Multiple(() =>
        {
            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.True);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        });
    }
    
    [Test]
    public void LoadSimpleDocFromString()
    {
        var tree = new RtfTree();

        var sr = new StreamReader(@"..\..\..\testdocs\testdoc1.rtf");
        var strDoc = sr.ReadToEnd();
        sr.Close();

        var res = tree.LoadRtfText(strDoc);

        sr = new StreamReader(@"..\..\..\testdocs\result1-1.txt");
        var strTree1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\result1-2.txt");
        var strTree2 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf1.txt");
        var rtf1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\text1.txt");
        var text1 = sr.ReadToEnd();
        sr.Close();

        Assert.That(res, Is.EqualTo(0));
        Assert.That(tree.MergeSpecialCharacters, Is.False);
        Assert.That(tree.ToString(), Is.EqualTo(strTree1));
        Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
        Assert.That(tree.Rtf, Is.EqualTo(rtf1));
        Assert.That(tree.Text, Is.EqualTo(text1));
    }
    
    [Test]
    public void LoadSimpleDocMergeSpecialFromString()
    {
        var tree = new RtfTree();
        tree.MergeSpecialCharacters = true;

        var sr = new StreamReader(@"..\..\..\testdocs\testdoc1.rtf");
        var strDoc = sr.ReadToEnd();
        sr.Close();

        var res = tree.LoadRtfText(strDoc);

        sr = new StreamReader(@"..\..\..\testdocs\result1-3.txt");
        var strTree1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\result1-4.txt");
        var strTree2 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\rtf1.txt");
        var rtf1 = sr.ReadToEnd();
        sr.Close();

        sr = new StreamReader(@"..\..\..\testdocs\text1.txt");
        var text1 = sr.ReadToEnd();
        sr.Close();

        Assert.That(res, Is.EqualTo(0));
        Assert.That(tree.MergeSpecialCharacters, Is.True);
        Assert.That(tree.ToString(), Is.EqualTo(strTree1));
        Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
        Assert.That(tree.Rtf, Is.EqualTo(rtf1));
        Assert.That(tree.Text, Is.EqualTo(text1));
    }
}