using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfBarTabTests {
    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("heading")]
    [InlineData("empty")]
    public void SaveAsPdf_EffectiveBarTabsPaintWithoutAnAuthoredTabAndReportApproximation(string story) {
        using WordDocument document = CreateDocument();
        WordParagraph paragraph = AddTarget(document, story);
        ConfigureParagraph(paragraph, story == "empty" ? string.Empty : "FIRST\nSECOND");
        SetTabs(paragraph, Tab(W.TabStopValues.Bar));
        string paragraphXml = paragraph._paragraph.OuterXml;
        string bodyXml = document.OpenXmlDocument.MainDocumentPart!.Document.Body!.OuterXml;
        string stylesXml = document.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "native-paragraph-bar-tab-approximated");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.DoesNotContain(result.Warnings, item => item.Code == "native-paragraph-bar-tab-omitted");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var bars = BarBounds(pdf);
        Assert.Equal(story == "empty" ? 1 : 2, bars.Length);
        Assert.All(bars, bar => Assert.True(bar.Height > 1D));
        Assert.All(bars, bar => Assert.Equal(bars[0].Left, bar.Left, 3));
        if (story is "body" or "heading" or "empty")
            Assert.All(bars, bar => Assert.Equal(108D, bar.Left, 2));
        if (story != "empty") {
            Assert.Contains("FIRST", string.Join("", pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
            Assert.Contains("SECOND", string.Join("", pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        }
        Assert.Equal(paragraphXml, paragraph._paragraph.OuterXml);
        Assert.Equal(bodyXml, document.OpenXmlDocument.MainDocumentPart.Document.Body!.OuterXml);
        Assert.Equal(stylesXml, document.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart.Styles!.OuterXml);
    }

    [Theory]
    [InlineData("direct", true)]
    [InlineData("style-masked", false)]
    [InlineData("style-inherited", true)]
    [InlineData("cleared", false)]
    public void SaveAsPdf_BarPaintAndLossUseTheSameNumberingOriginWinners(string origin, bool visible) {
        using WordDocument document = CreateDocument();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(72D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(54D);
        level.OpenXmlElement.PreviousParagraphProperties!.InsertAt(new W.Tabs(Tab(W.TabStopValues.Number)), 0);
        level.OpenXmlElement.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties(
            new W.RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
            new W.FontSize { Val = "24" });
        WordParagraph paragraph = list.AddItem("FIRST\nSECOND");
        ConfigureParagraph(paragraph, "FIRST\nSECOND");
        if (origin is "style-masked" or "style-inherited") {
            var styleProperties = new W.StyleParagraphProperties(new W.Tabs(Tab(W.TabStopValues.Bar)));
            var style = new W.Style(new W.StyleName { Val = "BarTabs" }, styleProperties) {
                Type = W.StyleValues.Paragraph, StyleId = "BarTabs"
            };
            document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            paragraph.SetStyleId("BarTabs");
            if (origin == "style-inherited") {
                W.NumberingProperties numbering = paragraph._paragraph.ParagraphProperties!.NumberingProperties!;
                numbering.Remove();
                styleProperties.AddChild(numbering, true);
            }
        } else SetTabs(paragraph, Tab(origin == "cleared" ? W.TabStopValues.Clear : W.TabStopValues.Bar));
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var bars = BarBounds(pdf);
        Assert.Equal(visible ? 2 : 0, bars.Length);
        if (visible) {
            // The rule stays at source position36pt even though the numbered
            // continuation text begins72pt from the containing text margin.
            Assert.All(bars, bar => Assert.Equal(108D, bar.Left, 2));
            Assert.Single(result.Warnings, item => item.Code == "native-paragraph-bar-tab-approximated");
        } else Assert.DoesNotContain(result.Warnings, item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_ImplicitParagraphAfterFinalTableRetainsDefaultBarWithoutChangingBody() {
        using WordDocument document = CreateDocument();
        SetDefaultBar(document);
        WordParagraph cell = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        ConfigureParagraph(cell, "CELL");
        SetTabs(cell, Tab(W.TabStopValues.Clear));
        W.Body body = document.OpenXmlDocument.MainDocumentPart!.Document.Body!;
        Assert.DoesNotContain(body.ChildElements, element => element is W.Paragraph);
        string bodyXml = body.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var bar = Assert.Single(BarBounds(pdf));
        var cellLetter = pdf.GetPage(1).Letters.First(letter => letter.Value == "C");
        Assert.True(bar.Top < cellLetter.StartBaseLine.Y,
            "The implicit paragraph's rule belongs below the table, not on its cell text line.");
        Assert.Equal(108D, bar.Left, 2);
        Assert.Single(result.Warnings, item => item.Code == "native-paragraph-bar-tab-approximated");
        Assert.Equal(bodyXml, body.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_EmptyReferencedFooterRetainsItsImplicitBarParagraph() {
        using WordDocument document = CreateDocument();
        SetDefaultBar(document);
        WordParagraph body = document.AddParagraph("BODY");
        SetTabs(body, Tab(W.TabStopValues.Clear));
        W.Footer footer = document.FooterDefaultOrCreate._footer!;
        footer.RemoveAllChildren();
        // Creating the footer also creates an empty default header. Isolate
        // this footer contract by explicitly clearing its inherited header bar.
        WordHeaderFooter header = document.HeaderDefaultOrCreate;
        header._header!.RemoveAllChildren();
        SetTabs(header.AddParagraph(), Tab(W.TabStopValues.Clear));
        string sourceXml = footer.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var bar = Assert.Single(BarBounds(pdf));
        Assert.Equal(108D, bar.Left, 2);
        Assert.True(bar.Top < 72D, "The implicit footer line remains inside its physical footer region.");
        Assert.Single(result.Warnings, item => item.Code == "native-paragraph-bar-tab-approximated");
        Assert.Equal(sourceXml, footer.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_FlatRunningProjectionReportsOmittedEffectiveBars() {
        using WordDocument document = CreateDocument();
        document.AddParagraph("BODY");
        WordParagraph paragraph = document.HeaderDefaultOrCreate.AddParagraph("FIRST");
        ConfigureParagraph(paragraph, "FIRST");
        SetTabs(paragraph, Tab(W.TabStopValues.Bar));
        WordPdfExportOptions options = Options();
        options.PdfOptions!.HeaderFontSize = 10D;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(options);
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "native-paragraph-bar-tab-omitted");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Empty(BarBounds(pdf));
        Assert.Contains("FIRST", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ObjectOnlyProjectionReportsOmissionBeforeAnyApproximation(bool header) {
        using WordDocument document = CreateDocument();
        document.AddParagraph("BODY");
        WordParagraph paragraph = header ? document.HeaderDefaultOrCreate.AddParagraph() : document.AddParagraph();
        ConfigureParagraph(paragraph, string.Empty);
        paragraph.AddShape(OfficeIMO.OfficeLength.Points(48D), OfficeIMO.OfficeLength.Points(36D), OfficeIMO.OfficeColor.Parse("#00A060"));
        SetTabs(paragraph, Tab(W.TabStopValues.Bar));
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal("native-paragraph-bar-tab-omitted", warning.Code);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Empty(BarBounds(pdf));
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsFilled);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ObjectOnlyCellRetainsItsDecoratedMark(bool header) {
        using WordDocument document = CreateDocument();
        document.AddParagraph("BODY");
        WordTable table = header ? document.HeaderDefaultOrCreate.AddTable(1, 1) : document.AddTable(1, 1);
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        ConfigureParagraph(paragraph, string.Empty);
        paragraph.AddShape(OfficeIMO.OfficeLength.Points(48D), OfficeIMO.OfficeLength.Points(36D), OfficeIMO.OfficeColor.Parse("#00A060"));
        SetTabs(paragraph, Tab(W.TabStopValues.Bar));
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal("native-paragraph-bar-tab-approximated", warning.Code);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Single(BarBounds(pdf));
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_MergeContinuationBarsDoNotCreatePaintOrLoss(bool vertical) {
        using WordDocument document = CreateDocument();
        WordTable table = document.AddTable(vertical ? 2 : 1, vertical ? 1 : 2);
        WordTableCell anchor = table.Rows[0].Cells[0];
        WordTableCell continuation = vertical ? table.Rows[1].Cells[0] : table.Rows[0].Cells[1];
        ConfigureParagraph(anchor.Paragraphs[0], "FIRST");
        ConfigureParagraph(continuation.Paragraphs[0], "OMITTED");
        SetTabs(continuation.Paragraphs[0], Tab(W.TabStopValues.Bar));
        if (vertical) {
            anchor.VerticalMerge = WordCellMerge.Restart;
            continuation.VerticalMerge = WordCellMerge.Continue;
        } else {
            anchor.HorizontalMerge = WordCellMerge.Restart;
            continuation.HorizontalMerge = WordCellMerge.Continue;
        }
        string tableXml = table._table.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        Assert.DoesNotContain(result.Warnings, item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Empty(BarBounds(pdf));
        Assert.Contains("FIRST", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("OMITTED", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Equal(tableXml, table._table.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_TableTextBoxBarReportsOmission() {
        using WordDocument document = CreateDocument();
        WordParagraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        var inner = new W.Paragraph(new W.ParagraphProperties(new W.Tabs(Tab(W.TabStopValues.Bar))),
            new W.Run(new W.Text("FIRST")));
        var shape = new DocumentFormat.OpenXml.Vml.Shape(new DocumentFormat.OpenXml.Vml.TextBox(new W.TextBoxContent(inner))) {
            Id = "CellBox", Type = "#_x0000_t202", Style = "width:96pt;height:36pt"
        };
        paragraph._paragraph.Append(new W.Run(new W.Picture(shape)));
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal("native-paragraph-bar-tab-omitted", warning.Code);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Empty(BarBounds(pdf));
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NestedTextBoxOwnersReportPaintedOrOmittedBars(bool surroundingText) {
        using WordDocument document = CreateDocument();
        WordParagraph outer = document.AddParagraph("PREFIX");
        WordParagraph middle = outer.AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single();
        ConfigureParagraph(middle, surroundingText ? "FIRST" : string.Empty);
        middle.AddTextBox("INNER", WordImageTextWrapping.Square);
        if (surroundingText) middle.AddText("SECOND");
        SetTabs(middle, Tab(W.TabStopValues.Bar));
        string sourceXml = outer._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal(surroundingText ? "native-paragraph-bar-tab-approximated" : "native-paragraph-bar-tab-omitted", warning.Code);
        Assert.Equal(surroundingText ? OfficeIMO.OfficeConversionLossKind.Approximation : OfficeIMO.OfficeConversionLossKind.Omission,
            warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(surroundingText ? 2 : 0, BarBounds(pdf).Length);
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("PREFIX", text, StringComparison.Ordinal);
        Assert.Contains("INNER", text, StringComparison.Ordinal);
        if (surroundingText) {
            Assert.True(text.IndexOf("FIRST", StringComparison.Ordinal) < text.IndexOf("INNER", StringComparison.Ordinal));
            Assert.True(text.IndexOf("INNER", StringComparison.Ordinal) < text.IndexOf("SECOND", StringComparison.Ordinal));
        }
        Assert.Equal(sourceXml, outer._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NestedRunningBoxOwnerReportsOmissionBeforeDeferredPainting(bool footer) {
        using WordDocument document = CreateDocument();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph outer = story.AddParagraph("PREFIX");
        WordParagraph middle = outer.AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single();
        ConfigureParagraph(middle, string.Empty);
        middle.AddTextBox("INNER", WordImageTextWrapping.Square);
        SetTabs(middle, Tab(W.TabStopValues.Bar));
        story.AddParagraph().AddField(WordFieldType.SectionPages);
        document.AddParagraph("BODY");
        string sourceXml = outer._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal("native-paragraph-bar-tab-omitted", warning.Code);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Empty(BarBounds(pdf));
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("PREFIX", text, StringComparison.Ordinal);
        Assert.Contains("INNER", text, StringComparison.Ordinal);
        Assert.Equal(sourceXml, outer._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("marker")]
    [InlineData("reference")]
    [InlineData("equation")]
    public void SaveAsPdf_NestedBoxOwnStructuralContentRetainsItsBar(string kind) {
        using WordDocument document = CreateDocument();
        WordParagraph outer = document.AddParagraph("PREFIX");
        WordParagraph middle = outer.AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single();
        ConfigureParagraph(middle, string.Empty);
        middle.AddTextBox("INNER", WordImageTextWrapping.Square);
        if (kind == "marker") {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
            list.Numbering.Levels[0].OpenXmlElement.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties(
                new W.RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
                new W.FontSize { Val = "24" });
            WordParagraph seed = list.AddItem("SEED");
            W.NumberingProperties numbering = seed._paragraph.ParagraphProperties!.NumberingProperties!;
            numbering.Remove();
            middle._paragraph.ParagraphProperties!.AddChild(numbering, true);
            seed._paragraph.Remove();
        } else if (kind == "reference") middle.AddFootNote("NOTEBODY");
        else middle.AddEquation("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>MATH</m:t></m:r></m:oMath>");
        SetTabs(middle, Tab(W.TabStopValues.Bar));
        string sourceXml = outer._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());

        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        Assert.Equal("native-paragraph-bar-tab-approximated", warning.Code);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Single(BarBounds(pdf));
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("INNER", text, StringComparison.Ordinal);
        Assert.Contains(kind == "marker" ? "1" : kind == "reference" ? "NOTEBODY" : "MATH", text, StringComparison.Ordinal);
        Assert.Equal(sourceXml, outer._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RunningMixedGroupBarReportMatchesItsActualOwnerLine(bool anchored) {
        using WordDocument document = CreateDocument();
        WordHeaderFooter story = document.HeaderDefaultOrCreate;
        WordParagraph outer = story.AddParagraph("PREFIX");
        WordParagraph middle = outer.AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single();
        ConfigureParagraph(middle, string.Empty);
        middle.AddTextBox("INNER", WordImageTextWrapping.Square);
        middle.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0D, 0D, 16D, 8D) { FillColorHex = "00FF00" },
            new WordShapeGroupItem(WordShapeType.Rectangle, 16D, 0D, 16D, 8D) { FillColorHex = "00FF00" }
        }, 180D, 30D);
        if (anchored) {
            var anchor = middle._paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>()
                .Single(candidate => candidate.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>()
                    .Any(data => data.Uri?.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"));
            anchor.BehindDoc = true;
            anchor.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare>();
            anchor.AddChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone(), true);
        }
        SetTabs(middle, Tab(W.TabStopValues.Bar));
        // Required field flow admits the group before ordinary running-story
        // drawing exclusions. The group itself is either anchored or flowed.
        story.AddParagraph().AddField(WordFieldType.SectionPages);
        document.AddParagraph("BODY");
        string sourceXml = story._header!.OuterXml;
        Assert.Empty(document.ValidateDocument());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(Options());
        // Capture the report before deferred running content is painted.
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code.StartsWith("native-paragraph-bar-tab-", StringComparison.Ordinal));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(anchored ? 1 : 0, BarBounds(pdf).Length);
        Assert.Contains(pdf.GetPage(1).Paths,
            path => path.IsFilled && path.FillColor.ToRGBValues() == (0D, 1D, 0D));
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("PREFIX", text, StringComparison.Ordinal);
        Assert.Contains("INNER", text, StringComparison.Ordinal);
        Assert.Equal(anchored ? "native-paragraph-bar-tab-approximated" : "native-paragraph-bar-tab-omitted", warning.Code);
        Assert.Equal(anchored ? OfficeIMO.OfficeConversionLossKind.Approximation : OfficeIMO.OfficeConversionLossKind.Omission,
            warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        Assert.Equal(sourceXml, story._header.OuterXml);
    }

    private static WordDocument CreateDocument() {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        return document;
    }

    private static WordPdfExportOptions Options() {
        var pdf = new PdfCreateOptions();
        pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(" ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.i".Select(character => (int)character).Concat(Enumerable.Range(32, 95)).Distinct().ToArray())));
        return new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = pdf
        };
    }

    private static WordParagraph AddTarget(WordDocument document, string story) {
        switch (story) {
            case "body":
            case "empty": return document.AddParagraph();
            case "table": return document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            case "header": document.AddParagraph("BODY"); return document.HeaderDefaultOrCreate.AddParagraph();
            case "footer": document.AddParagraph("BODY"); return document.FooterDefaultOrCreate.AddParagraph();
            case "heading":
                WordParagraph heading = document.AddParagraph();
                heading.Style = WordParagraphStyles.Heading1;
                return heading;
            default:
                WordParagraph reference = document.AddParagraph("REFERENCE");
                if (story == "endnote") reference.AddEndNote("NOTE");
                else reference.AddFootNote("NOTE");
                var main = document.OpenXmlDocument.MainDocumentPart!;
                W.FootnoteEndnoteType note = story == "endnote"
                    ? Assert.Single(main.EndnotesPart!.Endnotes!.Elements<W.Endnote>(), item => item.Id?.Value ==
                        Assert.Single(reference._paragraph.Descendants<W.EndnoteReference>()).Id!.Value)
                    : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), item => item.Id?.Value ==
                        Assert.Single(reference._paragraph.Descendants<W.FootnoteReference>()).Id!.Value);
                return new WordParagraph(document, Assert.Single(note.Elements<W.Paragraph>()));
        }
    }

    private static void ConfigureParagraph(WordParagraph paragraph, string text) {
        paragraph.Text = text;
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        W.ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        properties.SpacingBetweenLines = new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        };
    }

    private static W.TabStop Tab(W.TabStopValues alignment) => new() { Position = 720, Val = alignment };

    private static void SetTabs(WordParagraph paragraph, params W.TabStop[] tabs) {
        (paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties()).Tabs = new W.Tabs(tabs);
    }

    private static void SetDefaultBar(WordDocument document) {
        W.Styles styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.DocDefaults defaults = styles.DocDefaults ?? styles.PrependChild(new W.DocDefaults());
        W.ParagraphPropertiesDefault paragraph = defaults.ParagraphPropertiesDefault ?? defaults.AppendChild(new W.ParagraphPropertiesDefault());
        W.ParagraphPropertiesBaseStyle properties = paragraph.ParagraphPropertiesBaseStyle ?? paragraph.AppendChild(new W.ParagraphPropertiesBaseStyle());
        properties.AddChild(new W.Tabs(Tab(W.TabStopValues.Bar)), true);
    }

    private static UglyToad.PdfPig.Core.PdfRectangle[] BarBounds(PdfPigDocument pdf) => pdf.GetPages()
        .SelectMany(page => page.Paths).Where(path => path.IsStroked && Math.Abs(path.LineWidth - 0.14D) < .001D)
        .Select(path => path.GetBoundingRectangle()).Where(rectangle => rectangle is { Width: < .02D, Height: > 1D })
        .Select(rectangle => rectangle!.Value).ToArray();
}
