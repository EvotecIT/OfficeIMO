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
    [InlineData(false, false, false)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    public void PdfPictureBulletWarningOnlyIncludesRenderedHeaderFooterVariants(
        bool differentFirstPage,
        bool differentOddAndEvenPages,
        bool expectWarning) {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        WordHeader firstHeader = section.GetOrCreateHeader(WordHeaderFooterType.First);
        WordHeader evenHeader = section.GetOrCreateHeader(WordHeaderFooterType.Even);
        section.DifferentFirstPage = differentFirstPage;
        section.DifferentOddAndEvenPages = differentOddAndEvenPages;
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        AttachToList(firstHeader.AddParagraph("First picture marker"), list.NumberId);
        AttachToList(evenHeader.AddParagraph("Even picture marker"), list.NumberId);
        document.AddParagraph("Body");

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Equal(expectWarning, result.Warnings.Any(warning => warning.Code == "NativePictureBulletTextFallback"));
    }

    [Fact]
    public void HtmlHeaderListUsesListMarkupAndInactiveHeaderDiagnosticsAreSuppressed() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        AttachToList(document.Header!.Default!.AddParagraph("Header list item"), list.NumberId);
        document.AddParagraph("Body");

        string html = document.ToHtml(new WordToHtmlOptions { ExportHeadersAndFooters = true });
        Assert.Contains("<header", html, StringComparison.Ordinal);
        Assert.Contains("<ul", html, StringComparison.Ordinal);
        Assert.Contains("<li>Header list item</li>", html, StringComparison.Ordinal);

        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList pictureList = document.AddPictureBulletList(image, "Kulek.jpg");
        AttachToList(document.Header.Default.AddParagraph("Header picture marker"), pictureList.NumberId);
        var bodyOnly = document.ToHtmlResult(new WordToHtmlOptions { ExportHeadersAndFooters = false });
        Assert.DoesNotContain(bodyOnly.Report.Diagnostics, diagnostic => diagnostic.Code == "PictureBulletTextFallback");
    }

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
    public void ImageBodyListMarkerHonorsNumberingSuffixSpacing() {
        double nothingGap = RenderImageBodyListMarkerOffset(WordListLevelSuffix.Nothing);
        double spaceGap = RenderImageBodyListMarkerOffset(WordListLevelSuffix.Space);
        double tabGap = RenderImageBodyListMarkerOffset(WordListLevelSuffix.Tab);

        Assert.True(spaceGap > nothingGap + 1D, $"nothing={nothingGap}, space={spaceGap}");
        Assert.True(tabGap > spaceGap + 5D, $"space={spaceGap}, tab={tabGap}");
    }

    [Fact]
    public void ImageMarkerlessBodyListKeepsIndentAcrossSuffixes() {
        double nothingX = RenderImageMarkerlessBodyTextX(WordListLevelSuffix.Nothing);
        double spaceX = RenderImageMarkerlessBodyTextX(WordListLevelSuffix.Space);
        double tabX = RenderImageMarkerlessBodyTextX(WordListLevelSuffix.Tab);

        Assert.InRange(Math.Abs(nothingX - tabX), 0D, 0.01D);
        Assert.InRange(Math.Abs(spaceX - tabX), 0D, 0.01D);
    }

    [Fact]
    public void ImageTextBoxTabSuffixTargetsNumberingTextPosition() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 360;
        level.LevelSuffix = WordListLevelSuffix.Tab;
        WordTextBox box = document.AddTextBox("TabTextBox");
        box.WidthCentimeters = 12D;
        AttachToList(box.Paragraphs[0], list.NumberId);

        OfficeDrawingRichText rich = Assert.Single(document.CreateVisualSnapshot().Drawing.Elements
            .OfType<OfficeDrawingRichText>(), item => item.PlainText.Contains("TabTextBox", StringComparison.Ordinal));
        Assert.DoesNotContain('\t', rich.PlainText);
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            rich.Runs, rich.Width - rich.Padding.Horizontal, rich.Height - rich.Padding.Vertical, 1.25D,
            (text, size, _) => (text?.Length ?? 0) * size * 0.5D, wrap: true);
        OfficeRichTextLine line = Assert.Single(layout.Lines);
        int bodySegmentIndex = line.Segments.ToList().FindIndex(segment => segment.Text.Contains("TabTextBox", StringComparison.Ordinal));
        Assert.True(bodySegmentIndex > 0);
        double bodyOffset = line.OffsetX + line.Segments.Take(bodySegmentIndex).Sum(segment => segment.Width);
        Assert.InRange(bodyOffset, 85D, 95D);
    }

    [Fact]
    public void ListMarkerTraversalHandlesDeeplyNestedTextBoxesIteratively() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordTextBox current = document.AddTextBox("Level 0");
        for (int depth = 1; depth <= 128; depth++) {
            current = current.Paragraphs[0].AddTextBox("Level " + depth, WordImageTextWrapping.Square);
        }

        WordParagraph deepest = current.Paragraphs[0];
        AttachToList(deepest, list.NumberId);

        Dictionary<WordParagraph, (int Level, string Marker)> markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal("*", markers[deepest].Marker);
    }

    [Fact]
    public void ImageExportHandlesDeeplyNestedTextBoxesWithoutRecursiveTraversal() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordTextBox current = document.AddTextBox("Level 0");
        for (int depth = 1; depth <= 160; depth++) {
            current = current.Paragraphs[0].AddTextBox("Level " + depth, WordImageTextWrapping.Square);
        }

        AttachToList(current.Paragraphs[0], list.NumberId);
        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>(), rich =>
            rich.PlainText.Contains("Level 160", StringComparison.Ordinal) &&
            rich.PlainText.Contains("*", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfHeaderFooterRejectsExcessiveNestedTextBoxDepthWithoutStackOverflow() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordTextBox current = document.Header!.Default!.AddParagraph().AddTextBox("Level 0", WordImageTextWrapping.Square);
        AttachToList(current.Paragraphs[0], list.NumberId);
        for (int depth = 1; depth <= 20; depth++) {
            current = current.Paragraphs[0].AddTextBox("Level " + depth, WordImageTextWrapping.Square);
            AttachToList(current.Paragraphs[0], list.NumberId);
        }
        document.AddParagraph("Body");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => document.ToPdfBytes());
        Assert.Contains("text-box nesting exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
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
    public void ImageSplitTableRowPreservesVisibleListIndentAndHangingContinuation() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 5000U;
        section.PageSettings.Height = 2600U;
        section.SetMargins(WordMargin.Narrow);
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        list.Numbering.Levels[0].IndentationLeft = 1800;
        list.Numbering.Levels[0].IndentationHanging = 360;
        WordTable table = document.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 3600;
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new List<int> { 3600 };
        WordParagraph item = table.Rows[0].Cells[0].Paragraphs[0];
        item.Text = string.Join(" ", Enumerable.Range(1, 45).Select(index => "SplitVisible" + index.ToString("00")));
        AttachToList(item, list.NumberId);

        var allRichText = new List<(int Page, OfficeDrawingRichText Text)>();
        for (int pageIndex = 0; pageIndex < 8; pageIndex++) {
            WordDocumentVisualSnapshot page = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex });
            allRichText.AddRange(page.Drawing.Elements.OfType<OfficeDrawingRichText>().Select(text => (pageIndex, text)));
        }

        List<(int Page, OfficeDrawingRichText Text)> rendered = allRichText
            .Where(item => item.Text.PlainText.Contains("SplitVisible", StringComparison.Ordinal))
            .ToList();
        Assert.True(rendered.Select(item => item.Page).Distinct().Count() > 1,
            string.Join(" | ", rendered.Select(item => $"p{item.Page}:{item.Text.X}:{item.Text.PlainText}")));
        (int Page, OfficeDrawingRichText Text) first = rendered.OrderBy(item => item.Page).ThenBy(item => item.Text.Y).First();
        (int Page, OfficeDrawingRichText Text) continuation = rendered.First(item => item.Page > first.Page);
        (int Page, OfficeDrawingRichText Text) marker = Assert.Single(allRichText,
            item => item.Text.PlainText.Contains("*", StringComparison.Ordinal));
        Assert.Equal(first.Page, marker.Page);
        Assert.True(marker.Text.X > 50D, $"markerX={marker.Text.X}");
        Assert.True(first.Text.X > marker.Text.X + 10D,
            $"markerX={marker.Text.X}, firstX={first.Text.X}");
        Assert.InRange(Math.Abs(continuation.Text.X - first.Text.X), 0D, 0.01D);
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

    private static double RenderImageBodyListMarkerOffset(WordListLevelSuffix suffix) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        Level level = list.Numbering.Levels[0].OpenXmlElement;
        level.GetFirstChild<LevelSuffix>()?.Remove();
        level.Append(new LevelSuffix { Val = suffix.ToOpenXml() });
        WordParagraph item = document.AddParagraph("ImageBodySuffix");
        AttachToList(item, list.NumberId);

        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeDrawingText content = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "ImageBodySuffix");
        OfficeDrawingText marker = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "*");
        return content.X - marker.X;
    }

    private static double RenderImageMarkerlessBodyTextX(WordListLevelSuffix suffix) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 360;
        level.LevelSuffix = suffix;
        WordParagraph item = document.AddParagraph("MarkerlessBody");
        AttachToList(item, list.NumberId);

        return Assert.Single(document.CreateVisualSnapshot().Drawing.Elements.OfType<OfficeDrawingText>(),
            text => text.Text == "MarkerlessBody").X;
    }
}
