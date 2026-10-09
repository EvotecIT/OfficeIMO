using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentFlatStyleScopeTests {
    [Fact]
    public void CommonHeaderStyleRemainsDistinctFromBodyAutomaticStyleWithoutMutatingSource() {
        var source = OdtDocument.Create(); var body = source.AddParagraph("Body"); body.Bold = true;
        var header = source.PageLayout.Header.AddParagraph("Header"); header.Italic = true;
        var automatic = source.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!
            .Elements(OdfNamespaces.Style + "style").Single();
        automatic.SetAttributeValue(OdfNamespaces.Style + "name", "Shared"); body.StyleName = "Shared";
        var common = source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!
            .Elements(OdfNamespaces.Style + "style").Single();
        common.Remove(); common.SetAttributeValue(OdfNamespaces.Style + "name", "Shared");
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(common,
            new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "name", "Shared_flatCommon1"),
                new XAttribute(OdfNamespaces.Style + "family", "paragraph")));
        header.StyleName = "Shared";
        string[] before = new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()).ToArray();
        XDocument flat = source.ToFlatXml();
        Assert.Equal(before, new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()));
        string? bodyName = (string?)flat.Root!.Element(OdfNamespaces.Office + "body")!.Descendants(OdfNamespaces.Text + "p").Single().Attribute(OdfNamespaces.Text + "style-name");
        Assert.NotEqual("Shared", bodyName); Assert.NotEqual("Shared_flatCommon1", bodyName);
        Assert.Equal("Shared", (string?)flat.Descendants(OdfNamespaces.Style + "header").Single().Element(OdfNamespaces.Text + "p")!.Attribute(OdfNamespaces.Text + "style-name"));
        using var stream = new MemoryStream(); flat.Save(stream); stream.Position = 0;
        var read = OdtDocument.LoadFlatXml(stream);
        Assert.True(read.Paragraphs[0].Bold); Assert.NotEqual(true, read.Paragraphs[0].Italic);
        Assert.True(read.PageLayout.Header.Paragraphs[0].Italic); Assert.NotEqual(true, read.PageLayout.Header.Paragraphs[0].Bold);
    }

    [Fact]
    public void FlatMasterShapeNamesDoNotImportUnrelatedBodyDataStyles() {
        var source = OdgDocument.Create(); var page = source.AddPage();
        page.MasterShapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 2), "Header", "BodyDate");
        source.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(
            new XElement(OdfNamespaces.Number + "date-style", new XAttribute(OdfNamespaces.Style + "name", "BodyDate"), new XElement(OdfNamespaces.Number + "year")));
        using var flat = new MemoryStream(); source.SaveFlatXml(flat); flat.Position = 0; var read = OdgDocument.LoadFlatXml(flat);
        Assert.Null(read.Styles.FindDataStyle("BodyDate", "styles.xml")); Assert.NotNull(read.Styles.FindDataStyle("BodyDate"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlatConditionalDataStyleTargetsStayCommonWhenAutomaticNamesCollide(bool automaticCondition) {
        var source = ConditionalDataStyles(automaticCondition);
        string[] before = new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()).ToArray();
        XDocument flat = source.ToFlatXml();
        XElement conditional = flat.Descendants(OdfNamespaces.Number + "date-style").Single(element =>
            (string?)element.Attribute(OdfNamespaces.Style + "name") == "ConditionalDate");
        Assert.Equal("TargetDate", (string?)conditional.Element(OdfNamespaces.Style + "map")!.Attribute(OdfNamespaces.Style + "apply-style-name"));
        Assert.DoesNotContain(flat.Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Elements(), element =>
            (string?)element.Attribute(OdfNamespaces.Style + "name") == "TargetDate");
        using var stream = new MemoryStream(); flat.Save(stream); stream.Position = 0;
        var read = OdgDocument.LoadFlatXml(stream);
        AssertConditionalTarget(read, read.Pages[0]);
        Assert.Equal(before, new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedConditionalDataStyleTargetsStayCommonWhenAutomaticNamesCollide(bool automaticCondition) {
        var source = ConditionalDataStyles(automaticCondition);
        string[] before = new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()).ToArray();
        var destination = OdgDocument.Create();
        AssertConditionalTarget(destination, destination.ImportPage(source, 0));
        Assert.Equal(before, new[] { "content.xml", "styles.xml" }.Select(part => source.GetXml(part).ToString()));
    }

    [Fact]
    public void ImportRejectsAnAutomaticOnlyConditionalTargetWithoutChangingDestination() {
        var source = ConditionalDataStyles(false);
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Elements(OdfNamespaces.Number + "date-style")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == "TargetDate").Remove();
        var destination = OdgDocument.Create(); destination.AddPage("Existing");
        string[] before = new[] { "content.xml", "styles.xml" }.Select(part => destination.GetXml(part).ToString()).ToArray();
        Assert.Throws<InvalidDataException>(() => destination.ImportPage(source, 0));
        Assert.Equal(before, new[] { "content.xml", "styles.xml" }.Select(part => destination.GetXml(part).ToString()));
        Assert.Single(destination.Pages);
    }

    private static OdgDocument ConditionalDataStyles(bool automaticCondition) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var paragraph = page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 2), "").Paragraphs[0];
        paragraph.AddField(OdfTextFieldKind.Date, "conditional-cache").DataStyleName = "ConditionalDate";
        paragraph.AddField(OdfTextFieldKind.Date, "automatic-cache").DataStyleName = "TargetDate";
        XElement common = source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        common.Add(new XElement(OdfNamespaces.Number + "date-style", new XAttribute(OdfNamespaces.Style + "name", "TargetDate"), new XElement(OdfNamespaces.Number + "year")));
        foreach (string part in new[] { "content.xml", "styles.xml" })
            source.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(
                new XElement(OdfNamespaces.Number + "date-style", new XAttribute(OdfNamespaces.Style + "name", "TargetDate"), new XElement(OdfNamespaces.Number + "day")));
        (automaticCondition ? source.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")! : common).Add(
            new XElement(OdfNamespaces.Number + "date-style", new XAttribute(OdfNamespaces.Style + "name", "ConditionalDate"), new XElement(OdfNamespaces.Number + "month"),
                new XElement(OdfNamespaces.Style + "map", new XAttribute(OdfNamespaces.Style + "condition", "value()>0"), new XAttribute(OdfNamespaces.Style + "apply-style-name", "TargetDate"))));
        return source;
    }

    private static void AssertConditionalTarget(OdgDocument document, OdgPage page) {
        var fields = page.Shapes[0].Paragraphs[0].Fields.ToArray();
        OdfDataStyle conditional = document.Styles.FindDataStyle(fields[0].DataStyleName!)!;
        string targetName = (string)conditional.Element.Element(OdfNamespaces.Style + "map")!.Attribute(OdfNamespaces.Style + "apply-style-name")!;
        XElement target = document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Elements(OdfNamespaces.Number + "date-style")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == targetName);
        Assert.NotNull(target.Element(OdfNamespaces.Number + "year"));
        Assert.NotNull(document.Styles.FindDataStyle(fields[1].DataStyleName!)!.Element.Element(OdfNamespaces.Number + "day"));
        Assert.Equal(new[] { "conditional-cache", "automatic-cache" }, fields.Select(field => field.DisplayText));
    }

    [Fact]
    public void SharedFlatAutomaticStyleIsAvailableToBothBodyAndHeader() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        source.PageLayout.Header.AddParagraph("Header").StyleName = body.StyleName;
        XDocument flat = source.ToFlatXml();
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        OdtDocument loaded = OdtDocument.LoadFlatXml(stream);
        Assert.True(loaded.ContentBlocks.Single().Paragraph!.Bold);
        Assert.True(loaded.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(loaded.Validate().IsValid);
    }

    [Fact]
    public void ConflictingPartLocalNamesRemainDistinctAfterFlatRoundTrip() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        OdtParagraph header = source.PageLayout.Header.AddParagraph("Header");
        header.Italic = true;
        XElement headerStyle = source.Styles.FindInPart(OdfStyleFamily.Paragraph, header.StyleName!, "styles.xml")!.Element;
        headerStyle.SetAttributeValue(OdfNamespaces.Style + "name", body.StyleName);
        header.StyleName = body.StyleName;
        source.Package.MarkXmlDirty("styles.xml");
        XDocument flat = source.ToFlatXml();
        string[] names = flat.Root!.Element(OdfNamespaces.Office + "automatic-styles")!
            .Elements(OdfNamespaces.Style + "style").Select(element => (string)element.Attribute(OdfNamespaces.Style + "name")!).ToArray();
        Assert.Equal(names.Length, names.Distinct().Count());
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        OdtDocument loaded = OdtDocument.LoadFlatXml(stream);
        Assert.True(loaded.ContentBlocks.Single().Paragraph!.Bold);
        Assert.Null(loaded.ContentBlocks.Single().Paragraph!.Italic);
        Assert.True(loaded.PageLayout.Header.Paragraphs.Single().Italic);
        Assert.Null(loaded.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(loaded.Validate().IsValid);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlatStyleCollisionsPreserveFamilyAndCommonStyleReferences(bool commonTextStyle) {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        OdtParagraph header = source.PageLayout.Header.AddParagraph("Header ");
        header.Italic = true;
        OdtSpan span = header.AddSpan("Span");
        span.Bold = true;
        foreach (var pair in new[] {
            (OdfStyleFamily.Paragraph, body.StyleName!, "content.xml"),
            (OdfStyleFamily.Paragraph, header.StyleName!, "styles.xml"),
            (OdfStyleFamily.Text, span.StyleName!, "styles.xml") }) {
            source.Styles.FindInPart(pair.Item1, pair.Item2, pair.Item3)!.Element
                .SetAttributeValue(OdfNamespaces.Style + "name", "Shared");
        }
        if (commonTextStyle) {
            XElement textStyle = source.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!
                .Elements(OdfNamespaces.Style + "style").Single(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == "text");
            textStyle.Remove();
            source.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(textStyle);
        }
        body.StyleName = header.StyleName = span.StyleName = "Shared";
        source.Package.MarkXmlDirty("content.xml");
        source.Package.MarkXmlDirty("styles.xml");
        using var stream = new MemoryStream();
        source.ToFlatXml().Save(stream);
        stream.Position = 0;
        OdtDocument loaded = OdtDocument.LoadFlatXml(stream);
        Assert.True(loaded.ContentBlocks.Single().Paragraph!.Bold);
        OdtParagraph loadedHeader = loaded.PageLayout.Header.Paragraphs.Single();
        Assert.True(loadedHeader.Italic);
        Assert.True(loadedHeader.Spans.Single().Bold);
        Assert.True(loaded.Validate().IsValid);
    }

    [Fact]
    public void IdenticalStylesKeepTheirPartLocalDataStyleDependencies() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        OdtParagraph header = source.PageLayout.Header.AddParagraph("Header");
        header.Bold = true;
        foreach (var part in new[] { "content.xml", "styles.xml" }) {
            XElement automatic = source.Package.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
            XElement paragraph = automatic.Elements(OdfNamespaces.Style + "style").Single();
            paragraph.SetAttributeValue(OdfNamespaces.Style + "name", "Shared");
            paragraph.SetAttributeValue(OdfNamespaces.Style + "data-style-name", "Format");
            automatic.Add(new XElement(OdfNamespaces.Number + "number-style", new XAttribute(OdfNamespaces.Style + "name", "Format"),
                new XElement(OdfNamespaces.Number + "number", new XAttribute(OdfNamespaces.Number + "decimal-places", part == "content.xml" ? 1 : 2))));
            source.Package.MarkXmlDirty(part);
        }
        body.StyleName = header.StyleName = "Shared";
        XDocument flat = source.ToFlatXml();
        XElement definitions = flat.Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
        XElement masterParagraph = flat.Descendants(OdfNamespaces.Style + "header").Single().Element(OdfNamespaces.Text + "p")!;
        XElement masterStyle = definitions.Elements(OdfNamespaces.Style + "style").Single(element =>
            (string?)element.Attribute(OdfNamespaces.Style + "name") == (string?)masterParagraph.Attribute(OdfNamespaces.Text + "style-name"));
        XElement format = definitions.Elements(OdfNamespaces.Number + "number-style").Single(element =>
            (string?)element.Attribute(OdfNamespaces.Style + "name") == (string?)masterStyle.Attribute(OdfNamespaces.Style + "data-style-name"));
        Assert.Equal("2", (string?)format.Element(OdfNamespaces.Number + "number")!.Attribute(OdfNamespaces.Number + "decimal-places"));
    }

    [Fact]
    public void SparseReadIndexesTrackNativeAndMarkedXmlEditsWithoutExpandingRuns() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(1000, 1000).SetString("last");
        Assert.Equal("last", sheet.GetValue(1000, 1000).ToString());
        sheet.Cell(500, 500).SetString("middle");
        Assert.Equal("middle", sheet.GetValue(500, 500).ToString());
        Assert.Equal("last", sheet.GetValue(1000, 1000).ToString());
        sheet.Cell(500, 500).Formula = "of:=1+1";
        Assert.Equal("of:=1+1", sheet.GetFormula(500, 500));
        XElement firstRow = sheet.Element.Elements(OdfNamespaces.Table + "table-row").First();
        firstRow.SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", OdsRepeatModel.Read(firstRow, OdfNamespaces.Table + "number-rows-repeated") + 1);
        source.Package.MarkXmlDirty("content.xml");
        Assert.Equal("of:=1+1", sheet.GetFormula(501, 500));
        sheet.Cell(502, 500).SetString("after shift");
        Assert.Equal("of:=1+1", sheet.GetFormula(501, 500));
        Assert.Equal("after shift", sheet.GetValue(502, 500).ToString());
        Assert.True(sheet.RowRuns.Count < 10);
    }
}
