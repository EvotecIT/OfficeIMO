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

    [Fact]
    public void ADeclarationWrittenWithACharacterReferenceStillBindsItsPrefix() {
        const string xml = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/mai&#110;\"><p:cSld/></p:sld>";
        OfficeXmlTree tree = Parse(xml);
        tree.SetAttribute(tree.Root, OfficeXmlNamespaces.None, "show", "0");
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.PresentationML, "cSld"), tree.CreateElement(OfficeXmlNamespaces.PresentationML, "spTree"));

        Assert.Equal(xml.Replace("\"><p:cSld/>", "\" show=\"0\"><p:cSld><p:spTree/></p:cSld>"), Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void RejectsInsertionsThatWouldPutANodeInsideItself() {
        OfficeXmlTree tree = Parse("<r/>");
        int outer = tree.CreateElement(OfficeXmlNamespaces.None, "a");
        int inner = tree.CreateElement(OfficeXmlNamespaces.None, "b");
        tree.Append(outer, inner);

        Assert.Throws<InvalidOperationException>(() => tree.Append(outer, outer));
        tree.Remove(inner);
        tree.Append(tree.Root, inner);
        tree.Remove(inner);
        tree.Append(outer, inner);
        Assert.Throws<InvalidOperationException>(() => tree.Append(inner, outer));
    }

    [Fact]
    public void RejectsAReferenceNodeUnderAnotherParent() {
        OfficeXmlTree tree = Parse("<r><a/><b><x/><y/></b></r>");
        int a = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "a");
        int y = tree.FindChild(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "b"), OfficeXmlNamespaces.None, "y");

        Assert.Throws<InvalidOperationException>(() => tree.InsertBefore(a, tree.CreateElement(OfficeXmlNamespaces.None, "c"), y));
        Assert.Equal("<r><a/><b><x/><y/></b></r>", Encoding.UTF8.GetString(Write(tree)));
    }

    [Fact]
    public void DeclaresEachPrefixOncePerElement() {
        OfficeXmlTree tree = Parse("<r/>");
        tree.DeclareNamespace(tree.Root, "p", "urn:p");
        tree.DeclareNamespace(tree.Root, "p", "urn:p");
        tree.DeclareNamespace(tree.Root, string.Empty, string.Empty);
        tree.DeclareNamespace(tree.Root, string.Empty, string.Empty);

        Assert.Throws<InvalidOperationException>(() => tree.DeclareNamespace(tree.Root, "p", "urn:other"));
        Assert.Equal("<r xmlns:p=\"urn:p\" xmlns=\"\"/>", Encoding.UTF8.GetString(Write(tree)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChildrenKeepTheirNamespaceWhenAnAncestorDeclarationChanges(bool remove) {
        OfficeXmlTree tree = Parse("<r xmlns:p=\"urn:p\"><p:a p:k=\"v\"><p:b/></p:a></r>");

        if (remove) tree.RemoveAttribute(tree.Root, OfficeXmlNamespaces.Xmlns, "p");
        else tree.SetAttribute(tree.Root, OfficeXmlNamespaces.Xmlns, "p", "urn:other");

        XElement a = Load(Write(tree)).Root!.Elements().Single();
        Assert.Equal(XName.Get("a", "urn:p"), a.Name);
        Assert.Equal(XName.Get("b", "urn:p"), a.Elements().Single().Name);
        Assert.Equal("v", a.Attribute(XName.Get("k", "urn:p"))!.Value);
    }

    [Theory]
    [InlineData("<b/>")]
    [InlineData("<b xmlns:q=\"urn:other\" xmlns:t=\"urn:new\"/>")]
    public void AMovedElementKeepsTheBindingsOfItsEditedAttributesAndQNameValues(string destination) {
        OfficeXmlTree tree = Parse("<r><a xmlns:q=\"urn:q\" xmlns:t=\"urn:old\"><x q:k=\"old\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:type=\"t:W\"/></a>" + destination + "</r>");
        int x = tree.FirstChildOf(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "a"));
        tree.SetAttribute(x, tree.Namespaces.GetId("urn:q"), "k", "new");

        tree.Remove(x);
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "b"), x);

        XDocument written = XDocument.Load(new MemoryStream(Write(tree)));
        XElement moved = written.Root!.Element("b")!.Elements().Single();
        Assert.Equal("new", moved.Attribute(XName.Get("k", "urn:q"))!.Value);
        Assert.Equal("urn:old", moved.GetNamespaceOfPrefix("t")!.NamespaceName);
    }

    [Fact]
    public void EscapesTextThatAnEditBringsNextToOtherText() {
        OfficeXmlTree tree = Parse("<r>]]<x/>></r>");
        tree.Remove(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "x"));

        Assert.Equal("]]>", Load(Write(tree)).Root!.Value);
    }

    [Fact]
    public void WritesAndReadsDoublesWithXmlSchemaSpellings() {
        OfficeXmlTree tree = Parse("<r a=\"INF\" b=\"-INF\" c=\"Infinity\" d=\"1.5E3\"/>");

        Assert.Equal(double.PositiveInfinity, OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "a"));
        Assert.Equal(double.NegativeInfinity, OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "b"));
        Assert.Null(OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "c"));
        Assert.Equal(1500d, OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "d"));

        OfficeXmlViewValues.Write(tree, tree.Root, OfficeXmlNamespaces.None, "a", (double?)double.NegativeInfinity);
        OfficeXmlViewValues.Write(tree, tree.Root, OfficeXmlNamespaces.None, "b", (double?)double.PositiveInfinity);
        Assert.Equal("-INF", tree.GetAttribute(tree.Root, OfficeXmlNamespaces.None, "a"));
        Assert.Equal("INF", tree.GetAttribute(tree.Root, OfficeXmlNamespaces.None, "b"));
    }

    [Theory]
    [InlineData("<p:r xmlns:p=\"urn:old\"/>")]
    [InlineData("<r xmlns:p=\"urn:old\" p:a=\"1\"/>")]
    public void RejectsDeclarationEditsThatConflictWithNamesOnTheSameTag(string xml) {
        OfficeXmlTree tree = Parse(xml);

        Assert.Throws<InvalidOperationException>(() => tree.SetAttribute(tree.Root, OfficeXmlNamespaces.Xmlns, "p", "urn:new"));
        Assert.Throws<ArgumentException>(() => tree.SetAttribute(tree.Root, OfficeXmlNamespaces.Xmlns, "p", string.Empty));
        Assert.Equal(xml, Encoding.UTF8.GetString(Write(tree)));

        OfficeXmlTree nested = Parse("<r xmlns:p=\"urn:old\"><p:a/></r>");
        Assert.Throws<InvalidOperationException>(() => nested.DeclareNamespace(nested.FirstChildOf(nested.Root), "p", "urn:new"));
    }

    [Fact]
    public void AnElementMovedOutOfAMovedElementKeepsTheBindingsItWasWrittenWith() {
        OfficeXmlTree tree = Parse("<r xmlns:t=\"urn:old\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><outer><x xsi:type=\"t:W\"/></outer>"
            + "<d xmlns:t=\"urn:new\"/><n xmlns:t=\"urn:third\"/></r>");
        int outer = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "outer");
        int x = tree.FirstChildOf(outer);
        tree.Remove(outer);
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "d"), outer);

        tree.Remove(x);
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "n"), x);

        Assert.Equal("urn:old", MovedX(tree, "n").GetNamespaceOfPrefix("t")!.NamespaceName);
    }

    [Fact]
    public void AMoveAfterADeclarationEditKeepsEveryInheritedBinding() {
        OfficeXmlTree tree = Parse("<r><p xmlns:q=\"urn:q\" xmlns:t=\"urn:old\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><x q:k=\"v\" xsi:type=\"t:W\"/></p><d xmlns:t=\"urn:new\"/></r>");
        int parent = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "p");
        int x = tree.FirstChildOf(parent);
        tree.RemoveAttribute(parent, OfficeXmlNamespaces.Xmlns, "q");

        tree.Remove(x);
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "d"), x);

        XElement moved = MovedX(tree, "d");
        Assert.Equal("urn:old", moved.GetNamespaceOfPrefix("t")!.NamespaceName);
        Assert.Equal("v", moved.Attribute(XName.Get("k", "urn:q"))!.Value);
    }

    [Fact]
    public void ADetachedElementKeepsTheBindingsItHadWhenItWasRemoved() {
        OfficeXmlTree tree = Parse("<r><d xmlns:t=\"urn:old\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><x xsi:type=\"t:W\"/></d></r>");
        int parent = tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "d");
        int x = tree.FirstChildOf(parent);

        tree.Remove(x);
        tree.SetAttribute(parent, OfficeXmlNamespaces.Xmlns, "t", "urn:new");
        tree.Append(parent, x);

        Assert.Equal("urn:old", MovedX(tree, "d").GetNamespaceOfPrefix("t")!.NamespaceName);
    }

    [Fact]
    public void EscapesEveryTextPieceAfterOneThatIsWrittenAgain() {
        OfficeXmlTree tree = Parse("<r>A<x/><![CDATA[]]]]>>Z</r>");
        tree.Remove(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "x"));

        Assert.Equal("A]]>Z", Load(Write(tree)).Root!.Value);
    }

    [Fact]
    public void TrimsOnlyXmlWhitespaceAroundValues() {
        OfficeXmlTree tree = Parse("<r a=\" INF \" b=\" &#xA0;INF\" c=\"&#xA0;5\" d=\"\t7\r\n\"/>");

        Assert.Equal(double.PositiveInfinity, OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "a"));
        Assert.Null(OfficeXmlViewValues.ReadDouble(tree, tree.Root, OfficeXmlNamespaces.None, "b"));
        Assert.Null(OfficeXmlViewValues.ReadInt32(tree, tree.Root, OfficeXmlNamespaces.None, "c"));
        Assert.Equal(7, OfficeXmlViewValues.ReadInt32(tree, tree.Root, OfficeXmlNamespaces.None, "d"));
    }

    [Fact]
    public void ADeclarationOnAMovedElementKeepsTheBindingItCarriesForItsChildren() {
        OfficeXmlTree tree = Parse("<r><a xmlns:t=\"urn:old\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><m><x xsi:type=\"t:W\"/></m></a><d/></r>");
        int moved = tree.FirstChildOf(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "a"));
        tree.Remove(moved);
        tree.Append(tree.FindChild(tree.Root, OfficeXmlNamespaces.None, "d"), moved);

        tree.DeclareNamespace(moved, "t", "urn:new");

        Assert.Equal("urn:old", MovedX(tree, "m").GetNamespaceOfPrefix("t")!.NamespaceName);
    }

    [Fact]
    public void ADefaultDeclarationIgnoresUnprefixedAttributes() {
        OfficeXmlTree tree = Parse("<p:r xmlns:p=\"urn:p\" a=\"1\"/>");

        tree.DeclareNamespace(tree.Root, string.Empty, "urn:d");

        Assert.Equal("<p:r xmlns:p=\"urn:p\" a=\"1\" xmlns=\"urn:d\"/>", Encoding.UTF8.GetString(Write(tree)));
    }

    private static XElement MovedX(OfficeXmlTree tree, string parent) =>
        XDocument.Load(new MemoryStream(Write(tree))).Root!.Descendants(parent).Single().Element("x")!;

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
