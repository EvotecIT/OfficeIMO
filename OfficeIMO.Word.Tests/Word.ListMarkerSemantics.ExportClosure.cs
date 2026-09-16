using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordListMarkerSemanticsTests {
    [Theory]
    [InlineData(WordListLevelSuffix.Nothing, "*HeaderSuffix")]
    [InlineData(WordListLevelSuffix.Space, "* HeaderSuffix")]
    public void PdfHeaderListMarkerHonorsNumberingSuffix(WordListLevelSuffix suffix, string expected) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        Level level = list.Numbering.Levels[0].OpenXmlElement;
        level.GetFirstChild<LevelSuffix>()?.Remove();
        level.Append(new LevelSuffix { Val = suffix.ToOpenXml() });
        WordParagraph header = document.Header!.Default!.AddParagraph("HeaderSuffix");
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        string text = PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains(expected, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfHeaderListMarkerPreservesMarkerSpecificFormatting() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "FF0000", 18);
        list.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordParagraph header = document.Header!.Default!.AddParagraph("HeaderStyledMarker");
        header.FontSize = 9;
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        PdfTextSpan content = Assert.Single(spans, span => span.Text.Contains("HeaderStyledMarker", StringComparison.Ordinal));
        Assert.InRange(marker.FontSize, 17.5D, 18.5D);
        Assert.True(marker.IsBold);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), marker.Color);
        Assert.InRange(content.FontSize, 8.5D, 9.5D);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing, "*ImageSuffix")]
    [InlineData(WordListLevelSuffix.Space, "* ImageSuffix")]
    public void ImageTextBoxListMarkerHonorsNumberingSuffix(WordListLevelSuffix suffix, string expected) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        Level level = list.Numbering.Levels[0].OpenXmlElement;
        level.GetFirstChild<LevelSuffix>()?.Remove();
        level.Append(new LevelSuffix { Val = suffix.ToOpenXml() });
        WordTextBox box = document.AddTextBox("ImageSuffix");
        AttachToList(box.Paragraphs[0], list.NumberId);

        OfficeDrawingRichText rich = Assert.Single(document.CreateVisualSnapshot().Drawing.Elements
            .OfType<OfficeDrawingRichText>(), item => item.PlainText.Contains("ImageSuffix", StringComparison.Ordinal));
        Assert.Contains(expected, rich.PlainText, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureBulletFallbackIsReportedByHtmlAndImageExports() {
        using WordDocument document = WordDocument.Create();
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        list.AddItem("Picture item");

        var html = document.ToHtmlResult();
        Assert.Contains(html.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "PictureBulletTextFallback" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        Assert.Contains(snapshot.Diagnostics, diagnostic =>
            diagnostic.Code == WordImageExportDiagnosticCodes.LimitedPictureBulletTextFallback &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(document.ExportImage(OfficeImageExportFormat.Svg).Diagnostics, diagnostic =>
            diagnostic.Code == WordImageExportDiagnosticCodes.LimitedPictureBulletTextFallback);
    }

    [Fact]
    public void PictureBulletFallbackDiagnosticIsLimitedToTheRenderedPage() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Plain first page");
        document.AddPageBreak();
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        list.AddItem("Picture item on page two");

        WordDocumentVisualSnapshot firstPage = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = 0 });
        WordDocumentVisualSnapshot secondPage = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = 1 });
        Assert.DoesNotContain(firstPage.Diagnostics, diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.LimitedPictureBulletTextFallback);
        Assert.Contains(secondPage.Diagnostics, diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.LimitedPictureBulletTextFallback);
    }

    [Fact]
    public void PictureBulletFallbackInSplitTableRowIsReportedOnlyOnTheConsumingPage() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 5000U;
        section.PageSettings.Height = 3000U;
        section.SetMargins(WordMargin.Narrow);
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        WordTable table = document.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 3600;
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new List<int> { 3600 };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "PlainCell01 with enough content to occupy a line in a split table row.";
        for (int index = 2; index <= 6; index++) {
            cell.AddParagraph("PlainCell" + index.ToString("00") + " with enough content to occupy a line in a split table row.");
        }
        WordParagraph pictureItem = cell.AddParagraph("PictureCell07 on the later table-row fragment.");
        AttachToList(pictureItem, list.NumberId);

        WordDocumentVisualSnapshot[] pages = Enumerable.Range(0, 10)
            .Select(pageIndex => document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex }))
            .ToArray();
        int picturePageIndex = Array.FindIndex(pages, page => page.Drawing.Elements.OfType<OfficeDrawingRichText>().Any(text =>
            text.PlainText.Contains("PictureCell07", StringComparison.Ordinal)));
        Assert.True(picturePageIndex > 0,
            "Expected the picture-bullet paragraph on a later split-row page. " +
            string.Join(" || ", pages.Select((page, pageIndex) => pageIndex + ":" +
                string.Join(" | ", page.Drawing.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText)))));
        for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++) {
            bool hasDiagnostic = pages[pageIndex].Diagnostics.Any(diagnostic =>
                diagnostic.Code == WordImageExportDiagnosticCodes.LimitedPictureBulletTextFallback);
            Assert.Equal(pageIndex == picturePageIndex, hasDiagnostic);
        }
    }

    [Fact]
    public void TextBoxListIndentationIsKeptPerParagraphInPdfAndImageLayout() {
        using WordDocument document = WordDocument.Create();
        WordList bulletList = document.AddCustomBulletList('◆', "Arial", "000000");
        bulletList.Numbering.Levels[0].IndentationLeft = 1800;
        bulletList.Numbering.Levels[0].IndentationHanging = 360;
        WordList markerlessList = document.AddCustomList();
        markerlessList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        markerlessList.Numbering.Levels[0].IndentationLeft = 2160;
        WordTextBox box = document.AddTextBox("PlainBox");
        box.WidthCentimeters = 12D;
        box.Content!.Append(new Paragraph(new Run(new Text("BulletBox"))));
        box.Content.Append(new Paragraph(new Run(new Text("UnmarkedBox"))));
        AttachToList(box.Paragraphs[1], bulletList.NumberId);
        AttachToList(box.Paragraphs[2], markerlessList.NumberId);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        double plainX = Assert.Single(spans, span => span.Text.Contains("PlainBox", StringComparison.Ordinal)).X;
        Assert.True(spans.Any(span => span.Text.Contains("BulletBox", StringComparison.Ordinal)),
            PdfReadDocument.Open(document.ToPdfBytes()).ExtractText() + " | " +
            string.Join(" | ", spans.Select(span => $"'{span.Text}'@{span.X},{span.Y}+{span.Advance}")));
        double bulletX = Assert.Single(spans, span => span.Text.Contains("BulletBox", StringComparison.Ordinal)).X;
        double unmarkedX = Assert.Single(spans, span => span.Text.Contains("UnmarkedBox", StringComparison.Ordinal)).X;
        Assert.True(bulletX > plainX + 35D, $"plain={plainX}, bullet={bulletX}");
        Assert.True(unmarkedX > plainX + 55D, $"plain={plainX}, unmarked={unmarkedX}");

        OfficeDrawingRichText rich = Assert.Single(document.CreateVisualSnapshot().Drawing.Elements
            .OfType<OfficeDrawingRichText>(), item => item.PlainText.Contains("PlainBox", StringComparison.Ordinal));
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            rich.Runs, rich.Width - rich.Padding.Horizontal, rich.Height - rich.Padding.Vertical, 1.25D,
            (text, size, _) => (text?.Length ?? 0) * size * 0.5D, wrap: true);
        OfficeRichTextLine plainLine = Assert.Single(layout.Lines, line => line.Segments.Any(segment => segment.Text.Contains("PlainBox", StringComparison.Ordinal)));
        OfficeRichTextLine bulletLine = Assert.Single(layout.Lines, line => line.Segments.Any(segment => segment.Text.Contains("BulletBox", StringComparison.Ordinal)));
        OfficeRichTextLine unmarkedLine = Assert.Single(layout.Lines, line => line.Segments.Any(segment => segment.Text.Contains("UnmarkedBox", StringComparison.Ordinal)));
        Assert.True(bulletLine.OffsetX > plainLine.OffsetX + 40D,
            $"plain={plainLine.OffsetX}, bullet={bulletLine.OffsetX}, runs={string.Join(" | ", rich.Runs.Select(run => $"'{run.Text}':{run.ParagraphIndent?.FirstLineOffset}"))}");
        Assert.True(unmarkedLine.OffsetX > plainLine.OffsetX + 70D,
            $"plain={plainLine.OffsetX}, unmarked={unmarkedLine.OffsetX}");
    }

    [Fact]
    public void HostListMarkerStaysOutsideItsUnnumberedPdfTextBox() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordParagraph host = document.AddParagraph();
        WordTextBox box = host.AddTextBox("UnnumberedInner", WordImageTextWrapping.Square);
        AttachToList(host, list.NumberId);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text.Contains("1.", StringComparison.Ordinal));
        PdfTextSpan inner = Assert.Single(spans, span => span.Text.Contains("UnnumberedInner", StringComparison.Ordinal));
        Assert.NotEqual(marker.Y, inner.Y);
        Assert.Equal(1, PdfReadDocument.Open(document.ToPdfBytes()).ExtractText()
            .Split(new[] { "1." }, StringSplitOptions.None).Length - 1);
        Assert.NotNull(box);
    }

    [Fact]
    public void EmptyMarkerlessTextBoxDoesNotCreateAnEmptyRichTextBlock() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        WordTextBox box = document.AddTextBox(string.Empty);
        AttachToList(box.Paragraphs[0], list.NumberId);

        Assert.DoesNotContain(document.CreateVisualSnapshot().Drawing.Elements,
            element => element is OfficeDrawingRichText rich && rich.Runs.Count == 0);
        Assert.NotEmpty(document.ExportImage(OfficeImageExportFormat.Svg).Bytes);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing, 0D, 1D)]
    [InlineData(WordListLevelSuffix.Space, 2D, 5D)]
    [InlineData(WordListLevelSuffix.Tab, 20D, 40D)]
    public void PdfTableCellListMarkerHonorsNumberingSuffix(WordListLevelSuffix suffix, double minimumGap, double maximumGap) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        Level level = list.Numbering.Levels[0].OpenXmlElement;
        level.GetFirstChild<LevelSuffix>()?.Remove();
        level.Append(new LevelSuffix { Val = suffix.ToOpenXml() });
        WordTable table = document.AddTable(1, 1);
        WordParagraph cell = table.Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "SuffixCell";
        AttachToList(cell, list.NumberId);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        PdfTextSpan content = Assert.Single(spans, span => span.Text == "SuffixCell");
        Assert.InRange(content.X - marker.X - marker.Advance, minimumGap, maximumGap);
    }
}
