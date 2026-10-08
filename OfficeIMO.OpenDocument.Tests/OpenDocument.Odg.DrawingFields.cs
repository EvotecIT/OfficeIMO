using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgDrawingFieldTests {
    [Fact]
    public void NativeMasterPlaceholdersResolvePerPageAndDateTimeCachesSurviveBothContainers() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-master-fields.odg"));
        Assert.Equal(2, document.Pages.Count);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            for (int index = 0; index < read.Pages.Count; index++) {
                var page = read.Pages[index]; var result = page.ToDrawing();
                var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
                Assert.Contains("empty: " + (index + 1), text.PlainText);
                Assert.Contains("count: 2", text.PlainText); Assert.Contains("fixeddate: 10-1-2", text.PlainText);
                Assert.Contains("fixedtime: 03:04:00", text.PlainText);
                Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature.Contains(":field-") &&
                    mapping.Status is OdfConversionMappingStatus.Unsupported or OdfConversionMappingStatus.Skipped);
                Assert.Equal("<number>", page.MasterShapes[0].Paragraphs[0].Fields.Single().DisplayText);
                // Native automatic foreground opacity is a separate, retained typography limit.
                Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":text-opacity", StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Unsupported);
                Assert.Equal(before, Parts(read));
            }
        }
    }

    [Fact]
    public void SharedMasterFieldsFollowRenderedPageOrderCloneAndImportWithoutRefreshingSourceXml() {
        var document = OdgDocument.Create(); var first = document.AddPage("First"); var second = document.AddPage("Second");
        second.MasterPageName = first.MasterPageName;
        var paragraph = first.MasterShapes.AddTextBox(Bounds(), "Page ", "Header").Paragraphs[0];
        paragraph.AddField(OdfTextFieldKind.PageNumber, "stale").NumberFormat = "1";
        paragraph.AddText(" of "); paragraph.AddField(OdfTextFieldKind.PageCount, "99").NumberFormat = "1";
        document.ClonePage(0, "Copy"); document.MovePage(2, 0);
        foreach (var read in RoundTrips(document)) {
            string[] source = Parts(read);
            for (int index = 0; index < read.Pages.Count; index++) {
                Assert.Equal($"Page {index + 1} of 3", Text(read.Pages[index]));
                Assert.Equal(source, Parts(read));
            }
            var target = OdgDocument.Create(); target.AddPage("Empty"); target.ImportPage(read, 1, "Imported");
            Assert.Equal("Page 2 of 2", Text(target.Pages[1]));
            Assert.Equal(source, Parts(read));
            Assert.Equal("stale", read.Pages[1].MasterShapes[0].Paragraphs[0].Fields[0].DisplayText);
        }
    }

    [Fact]
    public void FieldsUseTheRenderedPageAcrossGroupsImagesAndLineCaptions() {
        var document = OdgDocument.Create(); document.AddPage(); var page = document.AddPage();
        var grouped = page.MasterShapes.AddGroup("Group").Children.AddRectangle(Bounds(), "Body");
        var raster = new OfficeRasterImage(2, 2); raster.Fill(OfficeColor.Parse("#2460a0"));
        var image = page.Shapes.AddImage(OfficePngWriter.Encode(raster), "caption.png", Bounds(), "Image");
        var line = page.Shapes.AddLine(OdfLength.Points(20), OdfLength.Points(400), OdfLength.Points(400), OdfLength.Points(400), "Line");
        foreach (var shape in new[] { grouped, image, line }) {
            var p = shape.AddParagraph(shape.Name + " "); p.FontFamily = "Arial"; p.FontSize = OdfLength.Points(12); p.Color = OdfColor.Parse("#000000");
            p.AddField(OdfTextFieldKind.PageNumber, "").NumberFormat = "1";
        }
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[1].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(new[] { "Body 2", "Image 2", "Line 2" }, result.Value.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText));
            Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>()); Assert.Equal(before, Parts(read));
        }
    }

    [Theory]
    [InlineData("1", false, "28")]
    [InlineData("a", false, "ab")]
    [InlineData("A", true, "BB")]
    [InlineData("i", false, "xxviii")]
    [InlineData("I", false, "XXVIII")]
    [InlineData("", false, "")]
    public void FieldNumberingFormatsAndLetterSynchronizationAreEditableAndSurviveBothContainers(string format, bool sync, string expected) {
        var document = OdgDocument.Create(); for (int index = 0; index < 28; index++) document.AddPage();
        var p = document.Pages[27].Shapes.AddTextBox(Bounds(), "[").Paragraphs[0];
        var field = p.AddField(OdfTextFieldKind.PageNumber, "cache"); field.NumberFormat = format; field.NumberLetterSync = sync; p.AddText("]");
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[27].Shapes[0].Paragraphs[0].Fields.Single();
            Assert.Equal(format, saved.NumberFormat); Assert.Equal(sync, saved.NumberLetterSync);
            Assert.Equal("[" + expected + "]", Text(read.Pages[27]));
        }
    }

    [Fact]
    public void PageNumberInheritsRenderedPageLayoutFormatAndPageCountUsesCurrentLogicalCount() {
        var document = OdgDocument.Create(); var first = document.AddPage(); var second = document.AddPage(); second.MasterPageName = first.MasterPageName;
        string layoutName = (string)first.Master!.Attribute(OdfNamespaces.Style + "page-layout-name")!;
        var layout = document.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "page-layout")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == layoutName).Element(OdfNamespaces.Style + "page-layout-properties")!;
        layout.SetAttributeValue(OdfNamespaces.Style + "num-format", "I"); document.MarkPartDirty("styles.xml");
        var p = first.MasterShapes.AddTextBox(Bounds(), "").Paragraphs[0];
        p.AddField(OdfTextFieldKind.PageNumber, "77"); p.AddText("/"); p.AddField(OdfTextFieldKind.PageCount, "999");
        Assert.Equal("II/2", Text(second));
        p.Fields[0].NumberFormat = "1"; Assert.Equal("2/2", Text(second));
        p.Fields[0].NumberFormat = null; Assert.Equal("II/2", Text(second));
    }

    [Fact]
    public void PageSelectionAndAdjustmentRespectSelectedAndAdjustedPageExistence() {
        var document = OdgDocument.Create(); for (int index = 0; index < 3; index++) document.AddPage();
        var first = document.Pages[0]; foreach (var page in document.Pages.Skip(1)) page.MasterPageName = first.MasterPageName;
        var p = first.MasterShapes.AddTextBox(Bounds(), "[").Paragraphs[0];
        var field = p.AddField(OdfTextFieldKind.PageNumber, "cache"); field.NumberFormat = "1";
        field.PageSelection = OdfTextFieldPageSelection.Next; field.PageAdjustment = -1; p.AddText("]");
        Assert.Equal(new[] { "[1]", "[2]", "[]" }, document.Pages.Select(Text));
        field.PageSelection = OdfTextFieldPageSelection.Previous; field.PageAdjustment = 1;
        Assert.Equal(new[] { "[]", "[2]", "[3]" }, document.Pages.Select(Text));
        field.PageSelection = OdfTextFieldPageSelection.Current; field.PageAdjustment = int.MaxValue;
        Assert.Equal("[]", Text(first));
        field.PageAdjustment = -1;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(new[] { "[]", "[1]", "[2]" }, read.Pages.Select(Text));
            Assert.Equal(-1, read.Pages[0].MasterShapes[0].Paragraphs[0].Fields.Single().PageAdjustment);
        }
    }

    [Fact]
    public void EmptyFieldCachesResolveBeforeWhitespaceAndRetainSpanAndHyperlinkStyles() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.MasterShapes.AddTextBox(Bounds(), "");
        var p = shape.Paragraphs[0]; p.FontSize = OdfLength.Points(14); p.FontFamily = "Liberation Sans";
        p.Element.Add(new XText("Page "));
        var span = p.AddRun(); span.Bold = true; span.AddField(OdfTextFieldKind.PageNumber, string.Empty).NumberFormat = "1";
        p.Element.Add(new XText(" of "));
        var link = p.AddHyperlink(string.Empty, "https://example.invalid/count"); link.Italic = true;
        link.AddField(OdfTextFieldKind.PageCount, string.Empty).NumberFormat = "1";
        p.Element.Add(new XText(" total")); document.MarkPartDirty("styles.xml");
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var text = Assert.Single(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Page 1 of 1 total", text.PlainText);
            Assert.True(text.Paragraphs[0].Runs.Single(run => run.Bold).Bold);
            var linked = text.Paragraphs[0].Runs.Single(run => run.Italic);
            Assert.Equal("1", linked.Text); Assert.True(linked.Italic); Assert.Equal(14, linked.FontSize);
            Assert.Contains("https://example.invalid/count", OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value)); Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void FixedNumbersAndDateTimeRetainSnapshotsWithExplicitReportsAndMissingCachesRejectStrictProjection() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var p = page.Shapes.AddTextBox(Bounds(), "", "Snapshots").Paragraphs[0];
        p.AddField(OdfTextFieldKind.PageNumber, "VII").IsFixed = true; p.AddText("/");
        p.AddField(OdfTextFieldKind.Date, "2010-01-02").IsFixed = true; p.AddText("/");
        p.AddField(OdfTextFieldKind.Time, "03:04");
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal("VII/2010-01-02/03:04", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        foreach (string feature in new[] { "field-fixed-cache", "field-date-time-cache" })
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":" + feature, StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Approximated);
        p.Fields[2].DisplayText = "";
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Contains(page.ToDrawing().Report.Mappings, mapping => mapping.Feature.EndsWith(":field-cache-missing", StringComparison.Ordinal));
        Assert.Throws<NotSupportedException>(() => p.Fields[1].NumberFormat = "1");
        Assert.Throws<NotSupportedException>(() => p.AddField(OdfTextFieldKind.PageCount).PageSelection = OdfTextFieldPageSelection.Next);
    }

    [Theory]
    [InlineData("num-format", "localized", true, "field-number-format")]
    [InlineData("num-letter-sync", "invalid", true, "field-number-format")]
    [InlineData("fixed", "invalid", false, "field-fixed")]
    [InlineData("select-page", "invalid", false, "field-page-selection")]
    [InlineData("page-adjust", "2147483648", false, "field-page-adjustment")]
    public void UnsupportedImportedFieldAttributesPreserveCacheAndRejectTheOmissionPolicy(string name, string value, bool style, string feature) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var p = page.Shapes.AddTextBox(Bounds(), "").Paragraphs[0];
        p.Element.Add(new XElement(OdfNamespaces.Text + "page-number", new XAttribute((style ? OdfNamespaces.Style : OdfNamespaces.Text) + name, value), "cache"));
        document.MarkPartDirty("content.xml"); string[] before = Parts(document);
        var result = page.ToDrawing(); Assert.Equal("cache", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":" + feature, StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FieldProjectionChargesSourceAndRenderedExpansionAcrossTheStory(bool expansion) {
        var document = OdgDocument.Create(); var page = document.AddPage(); document.AddPage();
        var shape = page.Shapes.AddRectangle(Bounds(), "Bounded"); var p = shape.AddParagraph();
        if (expansion) {
            p.AddText(new string('x', 100000 - 1));
            var field = p.AddField(OdfTextFieldKind.PageNumber, ""); // A supported two-character Roman value expands an empty source cache.
            field.NumberFormat = "I"; field.PageSelection = OdfTextFieldPageSelection.Next;
        } else {
            p.AddField(OdfTextFieldKind.PageNumber, new string('x', 100000));
            shape.AddParagraph("z");
        }
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()); Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "shape:Bounded:text" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void ScalarFieldCommentsCannotBypassTheStoryWideNodeLimit() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(Bounds(), "Nodes");
        for (int index = 0; index < 2; index++) {
            var field = new XElement(OdfNamespaces.Text + "page-number");
            for (int child = 0; child < 50000; child++) field.Add(new XComment("cache"));
            shape.AddParagraph().Element.Add(field);
        }
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()); Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("node limit"));
    }

    [Fact]
    public void AProjectedFieldCannotBypassTheContainerDepthLimitAfterAnEarlierVisibleParagraph() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(Bounds(), "Depth");
        shape.AddParagraph("Visible"); XElement parent = shape.AddParagraph().Element;
        for (int index = 0; index < 128; index++) { var span = new XElement(OdfNamespaces.Text + "span"); parent.Add(span); parent = span; }
        parent.Add(new XElement(OdfNamespaces.Text + "page-number"));
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("nesting limit"));
    }

    private static OdfRect Bounds() => OdfRect.FromCentimeters(1, 1, 14, 8);
    private static string Text(OdgPage page) => Assert.Single(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingRichText>()).PlainText;
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
