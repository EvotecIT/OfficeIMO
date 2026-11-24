using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// The editable XML tree writes an unchanged part back byte for byte, and a changed part as the same XML outside the
/// change, keeping prefixes, declarations, compatibility markup, comments and whitespace as they were written.
/// </summary>
public class OpenXmlXmlTreeTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void WritesEveryCorpusPartBackUnchangedAndRegeneratesItAsTheSameXml(string relativePath) {
        using OfficeZipPackageReader zip = OfficeZipPackageReader.Open(new MemoryStream(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath))));
        for (int index = 0; index < zip.Count; index++) {
            string name = zip.GetName(index);
            if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            Assert.True(zip.TryRead(name, out OfficePooledBytes content));
            byte[] source;
            using (content) source = content.Array.AsSpan(0, content.Length).ToArray();
            OfficeXmlTree tree = OfficeXmlTree.Parse(source, source.Length, new OfficeXmlNamespaces.Table());

            Assert.Equal(source, Write(tree));

            RegenerateEverything(tree, tree.Root);
            Assert.True(XNode.DeepEquals(Load(source), Load(Write(tree))), name + " changed when every start tag and text node was written again.");
        }
    }

    [Fact]
    public void KeepsCommentsCompatibilityMarkupAndUnknownAttributesAroundAnEdit() {
        const string xml = "<?xml version=\"1.0\"?>\n<!-- before -->\n<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:example\" mc:Ignorable=\"x\" x:extra='kept'>\n"
            + "  <!-- inside --><p:cSld name=\"A\"/>\n  <mc:AlternateContent><mc:Choice Requires=\"x\"><x:new/></mc:Choice><mc:Fallback><p:old/></mc:Fallback></mc:AlternateContent>\n</p:sld>\n<!-- after -->";
        OfficeXmlTree tree = Parse(xml);
        int slideData = tree.FindChild(tree.Root, OfficeXmlNamespaces.PresentationML, "cSld");

        tree.SetAttribute(slideData, OfficeXmlNamespaces.None, "name", "B & C");

        Assert.Equal(xml.Replace("name=\"A\"", "name=\"B &amp; C\""), Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void NewElementsUseThePrefixInScopeIncludingStrictNamespaces() {
        const string xml = "<s:sld xmlns:s=\"http://purl.oclc.org/ooxml/presentationml/main\"><s:cSld/></s:sld>";
        OfficeXmlTree tree = Parse(xml);
        int slideData = tree.FindChild(tree.Root, OfficeXmlNamespaces.PresentationML, "cSld");

        tree.Append(slideData, tree.CreateElement(OfficeXmlNamespaces.PresentationML, "spTree"));

        Assert.Equal("<s:sld xmlns:s=\"http://purl.oclc.org/ooxml/presentationml/main\"><s:cSld><s:spTree/></s:cSld></s:sld>", Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void DeclaresTheUsualPrefixWhereNoneIsInScope() {
        OfficeXmlTree tree = Parse("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>");
        int run = tree.CreateElement(OfficeXmlNamespaces.DrawingML, "t");
        tree.Append(tree.Root, run);
        tree.Append(run, tree.CreateText("x < y\r\n"));

        Assert.Equal("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:t xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">x &lt; y&#xD;\n</a:t></p:sld>",
            Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void RemovedNodesTakeTheCommentsWrittenBeforeThem() {
        OfficeXmlTree tree = Parse("<r><!-- a --><a/><!-- b --><b/></r>");
        tree.Remove(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "a"));

        Assert.Equal("<r><!-- b --><b/></r>", Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void InsertsChildrenInSchemaOrderWithChoicesSharingARank() {
        OfficeXmlTree tree = Parse("<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:br/><a:endParaRPr/></a:p>");
        var order = new OfficeXmlChildOrder(
            new[] { OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML, OfficeXmlNamespaces.DrawingML },
            new[] { "pPr", "r", "br", "endParaRPr" }, new[] { 0, 1, 1, 2 });

        tree.InsertOrdered(tree.Root, tree.CreateElement(OfficeXmlNamespaces.DrawingML, "r"), order);
        tree.GetOrInsertChild(tree.Root, OfficeXmlNamespaces.DrawingML, "pPr", order);

        Assert.Equal("<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr/><a:br/><a:r/><a:endParaRPr/></a:p>", Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void WritesAnEditedUtf16PartAsUtf8WithAMatchingDeclaration() {
        byte[] bytes = new UnicodeEncoding(bigEndian: false, byteOrderMark: true).GetPreamble()
            .Concat(Encoding.Unicode.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"yes\"?><r a=\"\u00E9\"/>")).ToArray();
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());

        tree.SetAttribute(tree.Root, OfficeXmlNamespaces.None, "b", "2");

        byte[] written = Write(tree);
        Assert.Equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><r a=\"\u00E9\" b=\"2\"/>", Encoding.UTF8.GetString(written));
        Assert.Equal("\u00E9", Load(written).Root!.Attribute("a")!.Value);
    }

    [Theory]
    [InlineData("<r><a xmlns:p=\"urn:x\"><p:x q:k=\"v\" xmlns:q=\"urn:q\"><p:y/></p:x></a><b/></r>")]
    [InlineData("<r xmlns:p=\"urn:other\"><a xmlns:p=\"urn:x\"><p:x q:k=\"v\" xmlns:q=\"urn:q\"><p:y/></p:x></a><b/></r>")]
    public void AMovedElementKeepsItsNamespacesAtItsNewPlace(string xml) {
        OfficeXmlTree tree = Parse(xml);
        int a = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "a");
        int b = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "b");
        int moved = tree.FirstChildOf(a);

        tree.Remove(moved);
        tree.Append(b, moved);

        XElement x = Load(Write(tree)).Root!.Element("b")!.Elements().Single();
        Assert.Equal(XName.Get("x", "urn:x"), x.Name);
        Assert.Equal(XName.Get("y", "urn:x"), x.Elements().Single().Name);
        Assert.Equal("v", x.Attribute(XName.Get("k", "urn:q"))!.Value);
    }

    private static OfficeXmlTree Parse(string xml) {
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        return OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
    }

    private static byte[] Write(OfficeXmlTree tree) {
        var output = new MemoryStream();
        var writer = new OfficeXmlUtf8Writer { Target = output };
        tree.WriteTo(writer);
        writer.Flush();
        return output.ToArray();
    }

    // Marks every start tag and text node changed without changing a value, so all of them are written again.
    private static void RegenerateEverything(OfficeXmlTree tree, int node) {
        if (tree.KindOf(node) == OfficeXmlTreeNodeKind.Text) {
            string text = tree.GetText(node);
            tree.SetText(node, text + "\u0001");
            tree.SetText(node, text);
            return;
        }
        tree.SetAttribute(node, OfficeXmlNamespaces.None, "regenerated", "1");
        tree.RemoveAttribute(node, OfficeXmlNamespaces.None, "regenerated");
        for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) RegenerateEverything(tree, child);
    }

    private static XDocument Load(byte[] data) {
        XDocument document = XDocument.Load(new MemoryStream(data), LoadOptions.PreserveWhitespace);
        foreach (XElement element in document.Descendants()) element.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
        return document;
    }
}
