using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentConversionContracts {
    [Fact]
    public void RejectingConversionResultAccessorsDisposeOwnedDocuments() {
        OdfConversionReport report = new OdfConversionReport("source", "target")
            .Add("unsupported", OdfConversionMappingStatus.Unsupported);
        var first = new DisposableProbe();
        var second = new DisposableProbe();

        Assert.Throws<OdfConversionLossException>(() =>
            new OdfConversionResult<DisposableProbe>(first, report).RequireNoLoss());
        Assert.Throws<OdfConversionLossException>(() =>
            new OdfConversionResult<DisposableProbe>(second, report)
                .GetValue(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));

        Assert.True(first.IsDisposed);
        Assert.True(second.IsDisposed);
    }

    [Fact]
    public void WordAndOdtRoundTripSemanticBlocksAndReportFidelity() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Native OpenDocument conversion").Style = WordParagraphStyles.Heading1;
        source.AddParagraph("Paragraph with structured content.").Bold = true;
        source.AddListNumbered().AddItem("First item");
        WordTable sourceTable = source.AddTable(2, 2);
        sourceTable.Rows[0].Cells[0].Paragraphs[0].Text = "A";
        sourceTable.Rows[1].Cells[1].Paragraphs[0].Text = "B";

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocumentResult();
        OdtDocument odt = toOdt.Value;
        Assert.True(odt.Validate().IsValid);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "headings");
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "tables");

        using var package = new MemoryStream(odt.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        Assert.Contains(reopened.ContentBlocks, block => block.Paragraph?.Text == "Native OpenDocument conversion");
        Assert.Contains(reopened.ContentBlocks, block => block.Table != null);

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocumentResult();
        using WordDocument roundTrip = toWord.Value;
        roundTrip.AddParagraph("Detached conversion remains editable");
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        WordDocumentSnapshot snapshot = roundTrip.CreateInspectionSnapshot();
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>(),
            paragraph => paragraph.Text.Contains("First item", StringComparison.Ordinal));
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements), block => block is WordTableSnapshot);
    }

    [Fact]
    public void WordAndOdtRoundTripCommonParagraphAndRunFormatting() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Styled text");
        paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
        paragraph.IndentationBeforePoints = 18;
        paragraph.IndentationAfterPoints = 9;
        paragraph.IndentationFirstLinePoints = 12;
        paragraph.LineSpacingBeforePoints = 6;
        paragraph.LineSpacingAfterPoints = 8;
        paragraph.ShadingFillColorHex = "EAF2F8";
        paragraph.Bold = true;
        paragraph.Italic = true;
        paragraph.Underline = WordUnderlineStyle.Single;
        paragraph.Strike = true;
        paragraph.FontSize = 13;
        paragraph.FontFamily = "Aptos";
        paragraph.ColorHex = "123456";
        paragraph.Highlight = WordHighlightColor.Yellow;

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocumentResult();
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(toOdt.Value.ToBytes()));
        OdtParagraph odtParagraph = Assert.Single(reopened.Paragraphs);
        OdtSpan odtSpan = Assert.Single(odtParagraph.Spans);

        Assert.Equal(OdtParagraphAlignment.Center, odtParagraph.Alignment);
        Assert.Equal(18, odtParagraph.IndentStart!.Value.ToPoints(), 3);
        Assert.Equal(9, odtParagraph.IndentEnd!.Value.ToPoints(), 3);
        Assert.Equal(12, odtParagraph.FirstLineIndent!.Value.ToPoints(), 3);
        Assert.Equal(6, odtParagraph.SpaceAbove!.Value.ToPoints(), 3);
        Assert.Equal(8, odtParagraph.SpaceBelow!.Value.ToPoints(), 3);
        Assert.Equal("#EAF2F8", odtParagraph.BackgroundColor!.Value.ToString());
        Assert.True(odtSpan.Bold);
        Assert.True(odtSpan.Italic);
        Assert.True(odtSpan.Underline);
        Assert.True(odtSpan.StrikeThrough);
        Assert.Equal("Aptos", odtSpan.FontFamily);
        Assert.Equal("#FFFF00", odtSpan.BackgroundColor!.Value.ToString());
        Assert.DoesNotContain(toOdt.Report.Mappings, mapping =>
            mapping.Feature == "paragraph-formatting" || mapping.Feature == "run-formatting");

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocumentResult();
        using WordDocument roundTrip = toWord.Value;
        WordParagraphSnapshot converted = Assert.Single(roundTrip.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());
        WordRunSnapshot run = Assert.Single(converted.Runs);

        Assert.Equal("Center", converted.Alignment);
        Assert.Equal(18, converted.IndentStartPoints!.Value, 3);
        Assert.Equal(9, converted.IndentEndPoints!.Value, 3);
        Assert.Equal(12, converted.IndentFirstLinePoints!.Value, 3);
        Assert.Equal(6, converted.SpaceAbovePoints!.Value, 3);
        Assert.Equal(8, converted.SpaceBelowPoints!.Value, 3);
        Assert.Equal("EAF2F8", converted.ShadingFillColorHex);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.True(run.Underline);
        Assert.True(run.Strike);
        Assert.Equal(13, run.FontSize);
        Assert.Equal("Aptos", run.FontFamily);
        Assert.Equal("123456", run.ColorHex);
        Assert.Equal("Yellow", run.HighlightColor);
    }

    [Fact]
    public void OdtInlineSyntaxPreservesMixedTextSpanAndHyperlinkOrderInWord() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.AddText("Before ");
        paragraph.AddSpan("bold").Bold = true;
        paragraph.AddText(" between ");
        paragraph.AddHyperlink("link", "https://example.com").Italic = true;
        paragraph.AddText(" after");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        OdtParagraph reopenedParagraph = Assert.Single(reopened.Paragraphs);
        Assert.Equal(
            new[] {
                OdtInlineNodeKind.Text,
                OdtInlineNodeKind.Span,
                OdtInlineNodeKind.Text,
                OdtInlineNodeKind.Hyperlink,
                OdtInlineNodeKind.Text
            },
            reopenedParagraph.InlineNodes.Select(node => node.Kind));

        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        WordParagraphSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());

        Assert.Equal("Before bold between link after", converted.Text);
        Assert.Equal(new[] { "Before ", "bold", " between ", "link", " after" },
            converted.Runs.Select(run => run.Text));
        Assert.True(converted.Runs[1].Bold);
        Assert.True(converted.Runs[3].Italic);
        Assert.Equal("https://example.com/", converted.Runs[3].HyperlinkUri);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting");
    }

    [Fact]
    public void OdtToWordDecodesPercentEncodedBookmarkLinks() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.AddBookmark("Section_1");
        paragraph.AddHyperlink("Jump", "#Section%5F1");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        WordParagraphSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());
        using var packageStream = new MemoryStream(target.ToBytes());
        using DocumentFormat.OpenXml.Packaging.WordprocessingDocument package =
            DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(packageStream, false);
        DocumentFormat.OpenXml.Wordprocessing.Hyperlink hyperlink = Assert.Single(
            package.MainDocumentPart!.Document!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>());

        Assert.Equal("Section_1", converted.BookmarkName);
        Assert.Equal("Section_1", hyperlink.Anchor!.Value);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdtInlineSyntaxPreservesImageOrderAndMapsBookmarksInWord() {
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.AddText("Before");
        paragraph.AddImage(png, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        paragraph.AddBookmark("target");
        paragraph.AddText("After");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(
            new[] { OdtInlineNodeKind.Text, OdtInlineNodeKind.Image, OdtInlineNodeKind.Bookmark, OdtInlineNodeKind.Text },
            Assert.Single(reopened.Paragraphs).InlineNodes.Select(node => node.Kind));

        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        WordParagraphSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());

        int imageIndex = converted.Runs.ToList().FindIndex(run => run.InlineImage != null);
        Assert.True(imageIndex > 0 && imageIndex < converted.Runs.Count - 1);
        Assert.Equal("Before", string.Concat(converted.Runs.Take(imageIndex).Select(run => run.Text)));
        Assert.Equal("After", string.Concat(converted.Runs.Skip(imageIndex + 1).Select(run => run.Text)));
        Assert.Equal("target", converted.BookmarkName);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "bookmarks" && mapping.Status == OdfConversionMappingStatus.Converted);
    }

    [Fact]
    public void ExcelAndOdsRoundTripTypedCellsFormulaMergeAndSparseLimits() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue("Amount").SetBold();
        sheet.CellAt(2, 1).SetValue(12.5m);
        sheet.CellAt(2, 2).SetFormula("SUM(A2:A2)");
        sheet.SetHyperlink(3, 1, "https://example.com", "Example");
        sheet.MergeRange("A4:B4");
        sheet.CellAt(4, 1).SetValue("Merged");
        source.SetNamedRange("Amounts", "'Data'!$A$2:$A$2", save: false);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        OdsDocument ods = toOds.Value;
        Assert.True(ods.Validate().IsValid);
        Assert.Equal(12.5m, ods.GetSheet("Data")!.Cell(1, 0).Value.AsDecimal());
        Assert.StartsWith("of:=", ods.GetSheet("Data")!.Cell(1, 1).Formula);
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "formulas" && mapping.Status == OdfConversionMappingStatus.Approximated);

        using var package = new MemoryStream(ods.ToBytes());
        OdsDocument reopened = OdsDocument.Load(package);
        OdfConversionResult<ExcelDocument> toExcel = reopened.ToExcelDocumentResult(new ExcelOpenDocumentConversionOptions {
            MaximumExpandedCells = 1000
        });
        using ExcelDocument roundTrip = toExcel.Value;
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        ExcelWorksheetSnapshot snapshot = Assert.Single(roundTrip.CreateInspectionSnapshot().Worksheets);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 1 && Convert.ToDecimal(cell.Value) == 12.5m);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 2 && cell.Formula != null);
        Assert.Contains(snapshot.MergedRanges, merge => merge.A1Range == "A4:B4");
    }

    [Fact]
    public void PowerPointAndOdpRoundTripSlidesShapesTablesNotesAndTransitions() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide();
        PowerPointTextBox title = slide.AddTextBoxPoints("OpenDocument deck", 24, 20, 240, 40);
        title.Paragraphs[0].Runs[0].Bold = true;
        slide.AddRectanglePoints(30, 80, 100, 50, "Panel").FillColor = "DDEEFF";
        PowerPointTable table = slide.AddTablePoints(2, 2, 150, 80, 160, 80);
        table.GetCell(0, 0).Text = "A";
        table.GetCell(1, 1).Text = "B";
        table.MergeCells(0, 0, 0, 1);
        slide.Notes.Text = "Speaker note";
        slide.Transition = PowerPointSlideTransition.Fade;

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocumentResult();
        OdpPresentation odp = toOdp.Value;
        Assert.True(odp.Validate().IsValid);
        OdpSlide odpSlide = Assert.Single(odp.Slides);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTextBox);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTable);
        Assert.Equal("Speaker note", Assert.Single(odpSlide.SpeakerNotes!.Paragraphs).Text);
        Assert.Contains(toOdp.Report.Mappings, mapping => mapping.Feature == "slide-transitions");

        using var package = new MemoryStream(odp.ToBytes());
        OdpPresentation reopened = OdpPresentation.Load(package);
        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopened.ToPowerPointPresentationResult();
        using PowerPointPresentation roundTrip = toPowerPoint.Value;
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        PowerPointSlide roundTripSlide = Assert.Single(roundTrip.Slides);
        Assert.Contains(roundTripSlide.TextBoxes, box => box.Text.Contains("OpenDocument deck", StringComparison.Ordinal));
        Assert.Single(roundTripSlide.Tables);
        Assert.Equal("Speaker note", roundTripSlide.GetSpeakerNotesText());
        Assert.Equal(PowerPointSlideTransition.Fade, roundTripSlide.Transition);
    }

    private sealed class DisposableProbe : IDisposable {
        internal bool IsDisposed { get; private set; }
        public void Dispose() => IsDisposed = true;
    }

    [Fact]
    public void PowerPointAndOdpRoundTripCommonRunFormatting() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun run = source.AddSlide().AddTextBoxPoints("Styled", 10, 10, 200, 40)
            .Paragraphs[0].Runs[0];
        run.Bold = true;
        run.Italic = true;
        run.Underline = true;
        run.Strikethrough = true;
        run.FontSize = 18;
        run.FontName = "Aptos";
        run.Color = "123456";
        run.HighlightColor = "FFF200";

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocumentResult();
        OdpRun odpRun = Assert.Single(Assert.IsType<OdpTextBox>(Assert.Single(toOdp.Value.Slides[0].Shapes))
            .Paragraphs[0].Runs);
        Assert.True(odpRun.Bold);
        Assert.True(odpRun.Italic);
        Assert.True(odpRun.Underline);
        Assert.True(odpRun.StrikeThrough);
        Assert.Equal("Aptos", odpRun.FontFamily);
        Assert.Equal("#123456", odpRun.Color!.Value.ToString());
        Assert.Equal("#FFF200", odpRun.BackgroundColor!.Value.ToString());
        Assert.DoesNotContain(toOdp.Report.Mappings, mapping => mapping.Feature == "run-format-details");

        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(toOdp.Value.ToBytes()));
        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopened.ToPowerPointPresentationResult();
        using PowerPointPresentation roundTrip = toPowerPoint.Value;
        PowerPointTextRun converted = roundTrip.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0];
        Assert.True(converted.Bold);
        Assert.True(converted.Italic);
        Assert.True(converted.Underline);
        Assert.True(converted.Strikethrough);
        Assert.Equal(18, converted.FontSize);
        Assert.Equal("Aptos", converted.FontName);
        Assert.Equal("123456", converted.Color);
        Assert.Equal("FFF200", converted.HighlightColor);
    }

    [Fact]
    public void PowerPointAndOdpTableCellsRoundTripParagraphsRunsAndHyperlinks() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTableCell sourceCell = source.AddSlide()
            .AddTablePoints(1, 1, 10, 10, 240, 80)
            .GetCell(0, 0);
        IReadOnlyList<PowerPointParagraph> sourceParagraphs = sourceCell.SetParagraphs(new[] { "Bold", "Second" });
        sourceParagraphs[0].Runs[0].Bold = true;
        PowerPointTextRun linkedRun = sourceParagraphs[0].AddRun(" link");
        linkedRun.Italic = true;
        linkedRun.SetHyperlink("https://example.com/table");
        sourceParagraphs[1].Runs[0].Underline = true;

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            });
        Assert.True(toOdp.Value.Validate().IsValid);
        OdpTableCell odpCell = Assert.IsType<OdpTable>(Assert.Single(toOdp.Value.Slides[0].Shapes)).Cell(0, 0);
        Assert.Equal(2, odpCell.Paragraphs.Count);
        Assert.Equal(new[] { "Bold", " link" }, odpCell.Paragraphs[0].InlineNodes.Select(node => node.Text));
        Assert.True(odpCell.Paragraphs[0].InlineNodes[0].Run!.Bold);
        Assert.True(odpCell.Paragraphs[0].InlineNodes[1].Hyperlink!.Italic);
        Assert.Equal("https://example.com/table", odpCell.Paragraphs[0].InlineNodes[1].Hyperlink!.Href);
        Assert.True(Assert.Single(odpCell.Paragraphs[1].Runs).Underline);

        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(toOdp.Value.ToBytes()));
        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopened.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            });
        using PowerPointPresentation roundTrip = toPowerPoint.Value;
        Assert.Empty(roundTrip.ValidateDocument());
        PowerPointTableCell roundTripCell = roundTrip.Slides[0].Tables.Single().GetCell(0, 0);
        Assert.Equal(2, roundTripCell.Paragraphs.Count);
        Assert.Equal(new[] { "Bold", " link" }, roundTripCell.Paragraphs[0].Runs.Select(run => run.Text));
        Assert.True(roundTripCell.Paragraphs[0].Runs[0].Bold);
        Assert.True(roundTripCell.Paragraphs[0].Runs[1].Italic);
        Assert.Equal("https://example.com/table", roundTripCell.Paragraphs[0].Runs[1].Hyperlink?.ToString());
        Assert.True(roundTripCell.Paragraphs[1].Runs[0].Underline);
    }

    [Fact]
    public void OdpInlineSyntaxPreservesMixedTextRunAndHyperlinkOrderInPowerPoint() {
        OdpPresentation source = OdpPresentation.Create();
        OdpParagraph paragraph = source.AddSlide("Inline").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 10, 3), null, "Text").AddParagraph();
        paragraph.AddText("Before ");
        paragraph.AddRun("bold").Bold = true;
        paragraph.AddText(" between ");
        paragraph.AddHyperlink("link", "https://example.com").Italic = true;
        paragraph.AddText(" after");

        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(source.ToBytes()));
        OdpParagraph reopenedParagraph = Assert.Single(Assert.IsType<OdpTextBox>(
            Assert.Single(reopened.Slides[0].Shapes)).Paragraphs);
        Assert.Equal(
            new[] {
                OdpInlineNodeKind.Text,
                OdpInlineNodeKind.Run,
                OdpInlineNodeKind.Text,
                OdpInlineNodeKind.Hyperlink,
                OdpInlineNodeKind.Text
            },
            reopenedParagraph.InlineNodes.Select(node => node.Kind));

        OdfConversionResult<PowerPointPresentation> conversion = reopened.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        var runs = target.Slides[0].TextBoxes.Single().Paragraphs[0].Runs;
        Assert.Equal(new[] { "Before ", "bold", " between ", "link", " after" }, runs.Select(run => run.Text));
        Assert.True(runs[1].Bold);
        Assert.True(runs[3].Italic);
        Assert.Equal("https://example.com/", runs[3].Hyperlink?.ToString());
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting");
    }

    [Fact]
    public void PowerPointRunHyperlinksMapToOdpHyperlinkSyntax() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun run = source.AddSlide().AddTextBoxPoints("Linked", 10, 10, 160, 40)
            .Paragraphs[0].Runs[0];
        run.Bold = true;
        run.SetHyperlink("https://example.com/path");

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpParagraph paragraph = Assert.Single(Assert.IsType<OdpTextBox>(
            Assert.Single(conversion.Value.Slides[0].Shapes)).Paragraphs);
        OdpInlineNode node = Assert.Single(paragraph.InlineNodes);

        Assert.Equal(OdpInlineNodeKind.Hyperlink, node.Kind);
        Assert.Equal("https://example.com/path", node.Hyperlink!.Href);
        Assert.True(node.Hyperlink.Bold);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "run-format-details");
    }

    [Fact]
    public void OdpToPowerPointRejectsTablesBeyondConfiguredBounds() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Bounded");
        slide.AddTable(OdfRect.FromCentimeters(1, 1, 10, 4), 1, 3);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToPowerPointPresentationResult(new PowerPointOpenDocumentConversionOptions {
                MaxTableRows = 2,
                MaxTableColumns = 2
            }));

        Assert.Contains("columns (3)", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointToOdpRejectsTablesBeyondConfiguredBounds() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide().AddTablePoints(1, 3, 10, 10, 120, 40);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToOpenDocumentResult(new PowerPointOpenDocumentConversionOptions {
                MaxTableRows = 2,
                MaxTableColumns = 2
            }));

        Assert.Contains("columns (3)", exception.Message, StringComparison.Ordinal);
    }
}
