using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class WordListMarkerSemanticsTests {
    [Fact]
    public void MarkdownPromotesChildrenOfTopLevelMarkerlessItemsWithoutPlaceholderMarkers() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        list.AddItem("Top level markerless", 0);
        list.AddItem("Visible child", 1);
        list.AddItem("Visible grandchild", 2);

        string markdown = document.ToMarkdown();
        Assert.Contains("Top level markerless", markdown, StringComparison.Ordinal);
        Assert.Contains("- Visible child", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(markdown.Split('\n'), line => line.Trim() == "-");
        Assert.DoesNotContain("  - Visible child", markdown, StringComparison.Ordinal);
        Assert.Contains("  - Visible grandchild", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectListLevelOutranksLinkedLevelWhenNumberIdIsInherited() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        for (int level = 0; level <= 3; level++) {
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        }
        list.Numbering.Levels[1].OpenXmlElement.Append(new ParagraphStyleIdInLevel { Val = "DirectLevelWins" });
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleParagraphProperties(new NumberingProperties(
            new NumberingId { Val = list.NumberId }))) {
            Type = StyleValues.Paragraph,
            StyleId = "DirectLevelWins"
        });
        WordParagraph item = document.AddParagraph("Direct level three");
        item._paragraph.ParagraphProperties = new ParagraphProperties(
            new ParagraphStyleId { Val = "DirectLevelWins" },
            new NumberingProperties(new NumberingLevelReference { Val = 3 }));

        Assert.Equal(3, WordDocumentTraversal.GetListInfo(item)!.Value.Level);
    }

    [Fact]
    public void OmittedNumberingSuffixUsesTabSemantics() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 360;
        level.OpenXmlElement.GetFirstChild<LevelSuffix>()?.Remove();
        WordParagraph item = document.AddParagraph("DefaultTabSuffix");
        AttachToList(item, list.NumberId);

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.Null(info.LevelSuffix);
        Assert.Equal("\t", WordDocumentTraversal.ResolveTextListMarkerSuffix(info.LevelSuffix));
        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeDrawingText marker = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "*");
        OfficeDrawingText content = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "DefaultTabSuffix");
        Assert.True(marker.Width > 10D, $"markerWidth={marker.Width}");
        Assert.InRange(Math.Abs(content.X - (marker.X + marker.Width)), 0D, 0.01D);
    }

    [Fact]
    public void ImageRightJustifiedNumberingSharesMarkerAndTextColumns() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelJustification = WordListLevelAlignment.Right;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        list.AddItem("Nine");
        list.AddItem("Ten");

        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeDrawingText nine = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "9.");
        OfficeDrawingText ten = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "10.");
        OfficeDrawingText nineText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Nine");
        OfficeDrawingText tenText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Ten");
        Assert.Equal(OfficeTextAlignment.Right, nine.Alignment);
        Assert.Equal(OfficeTextAlignment.Right, ten.Alignment);
        Assert.InRange(Math.Abs((nine.X + nine.Width) - (ten.X + ten.Width)), 0D, 0.01D);
        Assert.InRange(Math.Abs(nineText.X - tenText.X), 0D, 0.01D);
    }

    [Fact]
    public void PdfTableRightJustifiedNumberingSharesMarkerAndTextColumns() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelJustification = WordListLevelAlignment.Right;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        WordTable table = document.AddTable(2, 1);
        WordParagraph nineItem = table.Rows[0].Cells[0].Paragraphs[0];
        nineItem.Text = "TableNine";
        AttachToList(nineItem, list.NumberId);
        WordParagraph tenItem = table.Rows[1].Cells[0].Paragraphs[0];
        tenItem.Text = "TableTen";
        AttachToList(tenItem, list.NumberId);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan nine = Assert.Single(spans, span => span.Text == "9.");
        PdfTextSpan ten = Assert.Single(spans, span => span.Text == "10.");
        PdfTextSpan nineText = Assert.Single(spans, span => span.Text == "TableNine");
        PdfTextSpan tenText = Assert.Single(spans, span => span.Text == "TableTen");
        Assert.InRange(Math.Abs((nine.X + nine.Advance) - (ten.X + ten.Advance)), 0D, 1D);
        Assert.InRange(Math.Abs(nineText.X - tenText.X), 0D, 1D);
    }

    [Fact]
    public void PdfHeaderNestedTextBoxMarkerStylesFollowFlattenedTextOrder() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList nestedList = document.AddCustomBulletList('*', "Arial", "FF0000", 18);
        nestedList.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordList laterList = document.AddCustomBulletList('◆', "Arial", "0000FF", 10);
        WordList outerList = document.AddCustomBulletList('•', "Arial", "000000", 9);
        WordParagraph host = document.Header!.Default!.AddParagraph("Before ");
        WordTextBox outer = host.AddTextBox("Outer styled", WordImageTextWrapping.Square);
        host.AddText(" After");
        AttachToList(outer.Paragraphs[0], outerList.NumberId);
        WordTextBox nested = outer.Paragraphs[0].AddTextBox("Nested styled", WordImageTextWrapping.Square);
        AttachToList(nested.Paragraphs[0], nestedList.NumberId);
        outer.Content!.Append(new Paragraph(new Run(new Text("Later styled"))));
        WordParagraph later = outer.Paragraphs.Last();
        AttachToList(later, laterList.NumberId);
        document.AddParagraph("Body");

        Dictionary<WordParagraph, (int Level, string Marker)> markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.True(markers.TryGetValue(nested.Paragraphs[0], out var nestedResolved),
            string.Join(" | ", markers.Select(item => item.Key.Text + ":" + item.Value.Marker)));
        Assert.Equal("*", nestedResolved.Marker);
        Assert.True(markers.TryGetValue(later, out var laterResolved),
            string.Join(" | ", markers.Select(item => item.Key.Text + ":" + item.Value.Marker)));
        Assert.Equal("◆", laterResolved.Marker);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        string renderedSpans = string.Join(" | ", spans.Select(span => $"'{span.Text}'/{span.FontSize}/{span.Color}"));
        PdfTextSpan[] nestedMarkers = spans.Where(span => span.Text.Contains("*", StringComparison.Ordinal)).ToArray();
        PdfTextSpan[] laterMarkers = spans.Where(span => span.Text.Contains("◆", StringComparison.Ordinal)).ToArray();
        Assert.True(nestedMarkers.Length == 1, renderedSpans);
        Assert.True(laterMarkers.Length == 1, renderedSpans);
        PdfTextSpan nestedMarker = nestedMarkers[0];
        PdfTextSpan laterMarker = laterMarkers[0];
        Assert.InRange(nestedMarker.FontSize, 17.5D, 18.5D);
        Assert.True(nestedMarker.IsBold);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), nestedMarker.Color);
        Assert.InRange(laterMarker.FontSize, 9.5D, 10.5D);
        Assert.Equal(OfficeColor.FromRgb(0, 0, 255), laterMarker.Color);
        Assert.True(nestedMarker.Y > laterMarker.Y || nestedMarker.X < laterMarker.X);
    }

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

    [Fact]
    public void PdfHeaderTabSuffixUsesSupportedSpacingAndListIndentation() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        document.Header!.Default!.AddParagraph("PlainHeaderReference");
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 2160;
        level.IndentationHanging = 360;
        level.LevelSuffix = WordListLevelSuffix.Tab;
        WordParagraph header = document.Header.Default.AddParagraph("IndentedHeaderItem");
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan reference = Assert.Single(spans, span => span.Text.Contains("PlainHeaderReference", StringComparison.Ordinal));
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        PdfTextSpan content = Assert.Single(spans, span => span.Text.Contains("IndentedHeaderItem", StringComparison.Ordinal));
        Assert.True(marker.X > reference.X + 70D, $"reference={reference.X}, marker={marker.X}");
        Assert.InRange(Math.Abs(content.X - (marker.X + marker.Advance)), 0D, 0.5D);
        Assert.True(content.X > reference.X + 95D,
            $"marker={marker.X}+{marker.Advance}, content={content.X}");
    }

    [Theory]
    [InlineData(WordListLevelAlignment.Right)]
    [InlineData(WordListLevelAlignment.Center)]
    public void PdfHeaderNumberingHonorsLevelJustification(WordListLevelAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelJustification = alignment;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        NumberingSymbolRunProperties markerProperties = level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>() ??
            level.OpenXmlElement.AppendChild(new NumberingSymbolRunProperties());
        markerProperties.Append(new Bold(), new FontSize { Val = "36" });
        WordParagraph nineItem = document.Header!.Default!.AddParagraph("HeaderNine");
        nineItem.FontSize = 9;
        AttachToList(nineItem, list.NumberId);
        WordParagraph tenItem = document.Header.Default.AddParagraph("HeaderTen");
        tenItem.FontSize = 9;
        AttachToList(tenItem, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan nine = Assert.Single(spans, span => span.Text == "9.");
        PdfTextSpan ten = Assert.Single(spans, span => span.Text == "10.");
        PdfTextSpan nineText = Assert.Single(spans, span => span.Text.Contains("HeaderNine", StringComparison.Ordinal));
        PdfTextSpan tenText = Assert.Single(spans, span => span.Text.Contains("HeaderTen", StringComparison.Ordinal));
        double expectedStartShift = alignment == WordListLevelAlignment.Right ? 10D : 5D;
        Assert.InRange(nine.X - ten.X, expectedStartShift - 1D, expectedStartShift + 1D);
        Assert.True(Math.Abs(nineText.X - tenText.X) <= 1D,
            $"alignment={alignment}, nineText={nineText.X}, tenText={tenText.X}");
    }

    [Fact]
    public void PdfHeaderExtremeListIndentUsesBoundedMarkerSpacing() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = int.MaxValue;
        level.IndentationHanging = 0;
        level.LevelSuffix = WordListLevelSuffix.Tab;
        WordParagraph header = document.Header!.Default!.AddParagraph("ExtremeIndentHeader");
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        byte[] pdf = document.ToPdfBytes();

        Assert.InRange(pdf.Length, 1, 1_000_000);
        Assert.Contains("ExtremeIndentHeader", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfHeaderTextBoxPreservesInnerMarkerStyle() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "FF0000", 18);
        list.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordParagraph host = document.Header!.Default!.AddParagraph("Before ");
        WordTextBox box = host.AddTextBox("StyledHeaderBox", WordImageTextWrapping.Square);
        host.AddText(" After");
        AttachToList(box.Paragraphs[0], list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        Assert.InRange(marker.FontSize, 17.5D, 18.5D);
        Assert.True(marker.IsBold);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), marker.Color);
        string text = PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("Before * StyledHeaderBox", text, StringComparison.Ordinal);
        Assert.True(text.IndexOf("StyledHeaderBox", StringComparison.Ordinal) < text.IndexOf("After", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfHeaderRendersVisibleMarkerForEmptyListItem() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordParagraph emptyItem = document.Header!.Default!.AddParagraph();
        AttachToList(emptyItem, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan marker = Assert.Single(
            PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans(),
            span => span.Text == "*");
        Assert.True(marker.Advance > 0D);
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

    [Theory]
    [InlineData(WordListLevelAlignment.Right)]
    [InlineData(WordListLevelAlignment.Center)]
    public void TextBoxNumberingHonorsLevelJustificationInPdfAndImage(WordListLevelAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelJustification = alignment;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        NumberingSymbolRunProperties markerProperties = level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>() ??
            level.OpenXmlElement.AppendChild(new NumberingSymbolRunProperties());
        markerProperties.Append(new Bold(), new FontSize { Val = "36" });
        WordTextBox box = document.AddTextBox("TextBoxNine");
        box.WidthCentimeters = 12D;
        box.Content!.Append(new Paragraph(new Run(new Text("TextBoxTen"))));
        AttachToList(box.Paragraphs[0], list.NumberId);
        AttachToList(box.Paragraphs[1], list.NumberId);

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan pdfNine = Assert.Single(spans, span => span.Text == "9.");
        PdfTextSpan pdfTen = Assert.Single(spans, span => span.Text == "10.");
        PdfTextSpan pdfNineText = Assert.Single(spans, span => span.Text.Contains("TextBoxNine", StringComparison.Ordinal));
        PdfTextSpan pdfTenText = Assert.Single(spans, span => span.Text.Contains("TextBoxTen", StringComparison.Ordinal));
        double expectedStartShift = alignment == WordListLevelAlignment.Right ? 10D : 5D;
        Assert.InRange(pdfNine.X - pdfTen.X, expectedStartShift - 1D, expectedStartShift + 1D);
        Assert.InRange(Math.Abs(pdfNineText.X - pdfTenText.X), 0D, 1D);

        OfficeDrawingRichText rich = Assert.Single(document.CreateVisualSnapshot().Drawing.Elements
            .OfType<OfficeDrawingRichText>(), item => item.PlainText.Contains("TextBoxNine", StringComparison.Ordinal));
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
            rich.Runs, rich.Width - rich.Padding.Horizontal, rich.Height - rich.Padding.Vertical, 1.25D,
            MeasureRichTextWidth, wrap: true);
        OfficeRichTextLine nineLine = Assert.Single(layout.Lines, line => line.Segments.Any(segment => segment.Text.Contains("TextBoxNine", StringComparison.Ordinal)));
        OfficeRichTextLine tenLine = Assert.Single(layout.Lines, line => line.Segments.Any(segment => segment.Text.Contains("TextBoxTen", StringComparison.Ordinal)));
        (double NineX, double NineWidth) = GetRichTextTokenBounds(nineLine, "9.");
        (double TenX, double TenWidth) = GetRichTextTokenBounds(tenLine, "10.");
        double nineAnchor = alignment == WordListLevelAlignment.Right ? NineX + NineWidth : NineX + NineWidth / 2D;
        double tenAnchor = alignment == WordListLevelAlignment.Right ? TenX + TenWidth : TenX + TenWidth / 2D;
        Assert.True(Math.Abs(nineAnchor - tenAnchor) <= 2D,
            $"alignment={alignment}, nine={NineX}+{NineWidth}, ten={TenX}+{TenWidth}, runs={string.Join(" | ", rich.Runs.Select(run => $"'{run.Text}'/{run.ParagraphIndent?.FirstLineOffset}"))}");
        Assert.InRange(Math.Abs(
            GetRichTextTokenBounds(nineLine, "TextBoxNine").X -
            GetRichTextTokenBounds(tenLine, "TextBoxTen").X), 0D, 2D);
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
    public void ImageWrappedListContinuationKeepsAuthoredTextIndent() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 4200U;
        section.SetMargins(WordMargin.Narrow);
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        WordParagraph item = document.AddParagraph(string.Join(" ", Enumerable.Range(1, 8).Select(index => "WrapVisible" + index.ToString("00"))));
        AttachToList(item, list.NumberId);

        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeDrawingText marker = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "*");
        OfficeDrawingText body = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(),
            text => text.Text.StartsWith("Wrap", StringComparison.Ordinal));
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
            body.Text,
            body.Font.Size,
            body.Width - body.Padding.Horizontal,
            body.Height - body.Padding.Vertical,
            body.LineHeight.GetValueOrDefault(body.Font.Size * 1.25D) / body.Font.Size,
            1D,
            (text, size) => (text?.Length ?? 0) * size * 0.5D,
            wrap: true,
            paragraphIndent: body.ParagraphIndent);

        Assert.True(layout.Lines.Count > 1, string.Join(" | ", layout.Lines.Select(line => $"{line.Text}@{line.OffsetX}")));
        Assert.Equal(0D, layout.Lines[0].OffsetX);
        Assert.True(layout.Lines[1].OffsetX > 20D,
            $"marker={marker.X}, body={body.X}, continuationOffset={layout.Lines[1].OffsetX}");
        Assert.True(body.X + layout.Lines[1].OffsetX > marker.X + 30D,
            $"marker={marker.X}, continuation={body.X + layout.Lines[1].OffsetX}");
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
    public void ImageMarkerlessListPreservesExplicitZeroIndent() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("PlainZeroReference");
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        list.Numbering.Levels[0].IndentationLeft = 0;
        list.Numbering.Levels[0].IndentationHanging = 0;
        WordParagraph item = document.AddParagraph("MarkerlessZero");
        AttachToList(item, list.NumberId);

        OfficeDrawingText[] texts = document.CreateVisualSnapshot().Drawing.Elements.OfType<OfficeDrawingText>().ToArray();
        double referenceX = Assert.Single(texts, text => text.Text == "PlainZeroReference").X;
        double markerlessX = Assert.Single(texts, text => text.Text == "MarkerlessZero").X;
        Assert.InRange(Math.Abs(referenceX - markerlessX), 0D, 0.01D);
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

    [Theory]
    [InlineData(WordListLevelAlignment.Right)]
    [InlineData(WordListLevelAlignment.Center)]
    public void ImageSplitTableRowHonorsLevelJustification(WordListLevelAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 5000U;
        section.PageSettings.Height = 2400U;
        section.SetMargins(WordMargin.Narrow);
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1800;
        level.IndentationHanging = 720;
        level.LevelJustification = alignment;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        NumberingSymbolRunProperties markerProperties = level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>() ??
            level.OpenXmlElement.AppendChild(new NumberingSymbolRunProperties());
        markerProperties.Append(new Bold(), new FontSize { Val = "36" });
        WordTable table = document.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 3600;
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new List<int> { 3600 };
        WordTableCell cell = table.Rows[0].Cells[0];
        WordParagraph nine = cell.Paragraphs[0];
        nine.Text = "SplitNine";
        AttachToList(nine, list.NumberId);
        WordParagraph ten = cell.AddParagraph("SplitTen");
        AttachToList(ten, list.NumberId);
        for (int index = 0; index < 18; index++) {
            cell.AddParagraph("Following row content " + index.ToString("00"));
        }

        var rendered = new List<OfficeDrawingRichText>();
        for (int pageIndex = 0; pageIndex < 4; pageIndex++) {
            rendered.AddRange(document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex })
                .Drawing.Elements.OfType<OfficeDrawingRichText>());
        }

        OfficeDrawingRichText nineLine = Assert.Single(rendered, text => text.PlainText.StartsWith("9.", StringComparison.Ordinal));
        OfficeDrawingRichText tenLine = Assert.Single(rendered, text => text.PlainText.StartsWith("10.", StringComparison.Ordinal));
        OfficeRichTextRun nineMarker = nineLine.Runs[0];
        OfficeRichTextRun tenMarker = tenLine.Runs[0];
        double markerWidthDifference =
            MeasureRichTextWidth(tenMarker.Text, tenMarker.FontSize, tenMarker.FontFamily, tenMarker.FontStyle) -
            MeasureRichTextWidth(nineMarker.Text, nineMarker.FontSize, nineMarker.FontFamily, nineMarker.FontStyle);
        double expectedStartShift = alignment == WordListLevelAlignment.Right
            ? markerWidthDifference
            : markerWidthDifference / 2D;
        Assert.InRange(nineLine.X - tenLine.X, expectedStartShift - 1D, expectedStartShift + 1D);

        OfficeRichTextLine nineLayout = Assert.Single(OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
            nineLine.Runs, nineLine.Width - nineLine.Padding.Horizontal, nineLine.Height - nineLine.Padding.Vertical, 1.25D,
            MeasureRichTextWidth, wrap: false).Lines);
        OfficeRichTextLine tenLayout = Assert.Single(OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
            tenLine.Runs, tenLine.Width - tenLine.Padding.Horizontal, tenLine.Height - tenLine.Padding.Vertical, 1.25D,
            MeasureRichTextWidth, wrap: false).Lines);
        double nineTextX = nineLine.X + GetRichTextTokenBounds(nineLayout, "SplitNine").X;
        double tenTextX = tenLine.X + GetRichTextTokenBounds(tenLayout, "SplitTen").X;
        Assert.InRange(Math.Abs(nineTextX - tenTextX), 0D, 2D);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing, false)]
    [InlineData(WordListLevelSuffix.Space, false)]
    [InlineData(WordListLevelSuffix.Nothing, true)]
    [InlineData(WordListLevelSuffix.Space, true)]
    public void ImagePaginatedListLinesKeepContinuationIndent(WordListLevelSuffix suffix, bool richText) {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 5000U;
        section.PageSettings.Height = 2400U;
        section.SetMargins(WordMargin.Narrow);
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        list.Numbering.Levels[0].IndentationLeft = 1800;
        list.Numbering.Levels[0].IndentationHanging = 720;
        list.Numbering.Levels[0].LevelSuffix = suffix;
        string firstHalf = string.Join(" ", Enumerable.Range(0, 35).Select(index => "PageWord" + index.ToString("00"))) + " ";
        string secondHalf = string.Join(" ", Enumerable.Range(35, 35).Select(index => "PageWord" + index.ToString("00")));
        WordParagraph item = document.AddParagraph();
        item.Text = string.Empty;
        if (richText) {
            item.AddText(firstHalf);
            item.AddText(secondHalf).SetBold();
        } else {
            item.AddText(firstHalf + secondHalf);
        }
        AttachToList(item, list.NumberId);

        WordDocumentVisualSnapshot firstPage = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = 0 });
        if (richText) {
            OfficeDrawingRichText[] firstPageLines = firstPage.Drawing.Elements.OfType<OfficeDrawingRichText>()
                .Where(text => text.PlainText.Contains("PageWord", StringComparison.Ordinal))
                .ToArray();
            Assert.True(firstPageLines.Length >= 2, $"suffix={suffix}, expected at least two rich-text lines on the first page");
            OfficeDrawingRichText firstLine = firstPageLines[0];
            OfficeDrawingRichText firstPageContinuation = firstPageLines[1];
            OfficeDrawingRichText[] laterLines = Enumerable.Range(1, 6)
                .SelectMany(pageIndex => document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex })
                    .Drawing.Elements.OfType<OfficeDrawingRichText>())
                .Where(text => text.PlainText.Contains("PageWord", StringComparison.Ordinal))
                .ToArray();
            OfficeDrawingRichText continuation = laterLines.First();
            double firstTextX = firstLine.X + firstLine.Padding.Left + firstLine.ParagraphIndent.FirstLineOffset;
            double firstPageContinuationTextX = firstPageContinuation.X + firstPageContinuation.Padding.Left + firstPageContinuation.ParagraphIndent.FirstLineOffset;
            double continuationTextX = continuation.X + continuation.Padding.Left + continuation.ParagraphIndent.FirstLineOffset;
            Assert.True(firstPageContinuationTextX > firstTextX + 15D, $"suffix={suffix}, first={firstTextX}, same-page continuation={firstPageContinuationTextX}");
            Assert.True(continuationTextX > firstTextX + 15D, $"suffix={suffix}, first={firstTextX}, continuation={continuationTextX}");
        } else {
            OfficeDrawingText[] firstPageLines = firstPage.Drawing.Elements.OfType<OfficeDrawingText>()
                .Where(text => text.Text.Contains("PageWord", StringComparison.Ordinal))
                .ToArray();
            Assert.True(firstPageLines.Length >= 2, $"suffix={suffix}, expected at least two text lines on the first page");
            OfficeDrawingText firstLine = firstPageLines[0];
            OfficeDrawingText firstPageContinuation = firstPageLines[1];
            OfficeDrawingText continuation = Enumerable.Range(1, 6)
                .SelectMany(pageIndex => document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex })
                    .Drawing.Elements.OfType<OfficeDrawingText>())
                .First(text => text.Text.Contains("PageWord", StringComparison.Ordinal));
            double firstTextX = firstLine.X + firstLine.Padding.Left + firstLine.ParagraphIndent.FirstLineOffset;
            double firstPageContinuationTextX = firstPageContinuation.X + firstPageContinuation.Padding.Left + firstPageContinuation.ParagraphIndent.FirstLineOffset;
            double continuationTextX = continuation.X + continuation.Padding.Left + continuation.ParagraphIndent.FirstLineOffset;
            Assert.True(firstPageContinuationTextX > firstTextX + 15D, $"suffix={suffix}, first={firstTextX}, same-page continuation={firstPageContinuationTextX}");
            Assert.True(continuationTextX > firstTextX + 15D, $"suffix={suffix}, first={firstTextX}, continuation={continuationTextX}");
        }
    }

    [Fact]
    public void ImageTextBoxTabSuffixBoundsSynthesizedSpacing() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        list.Numbering.Levels[0].IndentationLeft = int.MaxValue;
        list.Numbering.Levels[0].IndentationHanging = int.MaxValue;
        list.Numbering.Levels[0].LevelSuffix = WordListLevelSuffix.Tab;
        WordTextBox box = document.AddTextBox("BoundedTabSpacer");
        box.WidthCentimeters = 10_000D;
        AttachToList(box.Paragraphs[0], list.NumberId);

        OfficeDrawingRichText rich = Assert.Single(document.CreateVisualSnapshot().Drawing.Elements
            .OfType<OfficeDrawingRichText>(), text => text.PlainText.Contains("BoundedTabSpacer", StringComparison.Ordinal));
        OfficeRichTextRun marker = Assert.Single(rich.Runs, run => run.Text.StartsWith("*", StringComparison.Ordinal));
        Assert.InRange(marker.Text.Length, 2, 8_193);
    }

    [Fact]
    public void ImageSplitTableRowKeepsPostImageListTextAtContinuationIndent() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 5000U;
        section.PageSettings.Height = 2400U;
        section.SetMargins(WordMargin.Narrow);
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        list.Numbering.Levels[0].IndentationLeft = 1800;
        list.Numbering.Levels[0].IndentationHanging = 720;
        list.Numbering.Levels[0].LevelSuffix = WordListLevelSuffix.Tab;
        WordTable table = document.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 3600;
        table.ColumnWidthType = WordTableWidthUnit.Dxa;
        table.ColumnWidth = new List<int> { 3600 };
        WordTableCell cell = table.Rows[0].Cells[0];
        WordParagraph item = cell.Paragraphs[0];
        item.Text = string.Empty;
        item.AddText("BeforeInlineImage");
        item.AddText(string.Empty).InsertImage(
            Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"), 12, 12,
            WordImageTextWrapping.InLineWithText, "Inline continuation marker");
        item.AddText("AfterInlineImage");
        AttachToList(item, list.NumberId);
        for (int index = 0; index < 18; index++) {
            cell.AddParagraph("Following row content " + index.ToString("00"));
        }

        var rendered = new List<OfficeDrawingRichText>();
        for (int pageIndex = 0; pageIndex < 6; pageIndex++) {
            rendered.AddRange(document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = pageIndex })
                .Drawing.Elements.OfType<OfficeDrawingRichText>());
        }

        string renderedText = string.Join(" | ", rendered.Select(text => $"'{text.PlainText}'@{text.X}"));
        OfficeDrawingRichText marker = Assert.Single(rendered, text => text.PlainText.StartsWith("*", StringComparison.Ordinal));
        OfficeDrawingRichText before = Assert.Single(rendered, text => text.PlainText.StartsWith("Before", StringComparison.Ordinal));
        OfficeDrawingRichText after = Assert.Single(rendered, text => text.PlainText.StartsWith("After", StringComparison.Ordinal));
        Assert.InRange(Math.Abs(before.X - after.X), 0D, 0.01D);
        Assert.True(after.X > marker.X + 25D, $"marker={marker.X}, before={before.X}, after={after.X}; {renderedText}");
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
    [InlineData(WordListLevelSuffix.Tab, 10D, 15D)]
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

    private static double MeasureRichTextWidth(string? text, double fontSize, string? fontFamily, OfficeFontStyle style) {
        var font = new OfficeFontInfo(fontFamily ?? OfficeFontInfo.Default.FamilyName, fontSize, style);
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
        return measurer.MeasureWidth(text, measurer.CreateStyle(font, 72D));
    }

    private static (double X, double Width) GetRichTextTokenBounds(OfficeRichTextLine line, string token) {
        double cursor = line.OffsetX;
        foreach (OfficeRichTextSegment segment in line.Segments) {
            int tokenIndex = segment.Text.IndexOf(token, StringComparison.Ordinal);
            if (tokenIndex >= 0) {
                OfficeFontStyle style = OfficeFontStyle.Regular;
                if (segment.Bold) style |= OfficeFontStyle.Bold;
                if (segment.Italic) style |= OfficeFontStyle.Italic;
                double prefixWidth = MeasureRichTextWidth(segment.Text.Substring(0, tokenIndex), segment.FontSize, segment.FontFamily, style);
                double tokenWidth = MeasureRichTextWidth(token, segment.FontSize, segment.FontFamily, style);
                return (cursor + prefixWidth, tokenWidth);
            }

            cursor += segment.Width;
        }

        throw new Xunit.Sdk.XunitException("Token was not present in the rich-text line: " + token);
    }
}
