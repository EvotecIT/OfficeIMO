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

        string markdown = document.ToMarkdown().ToMarkdown();
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
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(18D);
        level.OpenXmlElement.GetFirstChild<LevelSuffix>()?.Remove();
        WordParagraph item = document.AddParagraph("DefaultTabSuffix");
        AttachToList(item, list.NumberId);

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.Null(info.LevelSuffix);
        Assert.Equal("\t", WordDocumentTraversal.ResolveTextListMarkerSuffix(info.LevelSuffix));
    }

    [Fact]
    public void PdfTableRightJustifiedNumberingSharesMarkerAndTextColumns() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(36D);
        level.LevelJustification = WordListLevelAlignment.Right;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        WordTable table = document.AddTable(2, 1);
        WordParagraph nineItem = table.Rows[0].Cells[0].Paragraphs[0];
        nineItem.Text = "TableNine";
        AttachToList(nineItem, list.NumberId);
        WordParagraph tenItem = table.Rows[1].Cells[0].Paragraphs[0];
        tenItem.Text = "TableTen";
        AttachToList(tenItem, list.NumberId);

        byte[] pdf = document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        (double Start, double End) nine = PdfLetterRun(pdf, "9.");
        (double Start, double End) ten = PdfLetterRun(pdf, "10.");
        (double Start, double End) nineText = PdfLetterRun(pdf, "TableNine");
        (double Start, double End) tenText = PdfLetterRun(pdf, "TableTen");
        Assert.InRange(Math.Abs(nine.End - ten.End), 0D, 1D);
        Assert.InRange(Math.Abs(nineText.Start - tenText.Start), 0D, 1D);
    }

    [Fact]
    public void PdfHeaderNestedTextBoxMarkerStylesFollowFlattenedTextOrder() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList nestedList = document.AddCustomBulletList('*', "Arial", "FF0000", global::OfficeIMO.OfficeLength.Points(18));
        nestedList.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordList laterList = document.AddCustomBulletList('◆', "Arial", "0000FF", global::OfficeIMO.OfficeLength.Points(10));
        WordList outerList = document.AddCustomBulletList('•', "Arial", "000000", global::OfficeIMO.OfficeLength.Points(9));
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

        PdfTextSpan[] spans = PdfReadDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).Pages[0].GetTextSpans().ToArray();
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

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.Equal(expectWarning, result.Warnings.Any(warning => warning.Code == "native-picture-bullet-text-fallback"));
    }

    [Fact]
    public void HtmlHeaderListUsesListMarkupAndInactiveHeaderDiagnosticsAreSuppressed() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        AttachToList(document.Header!.Default!.AddParagraph("Header list item"), list.NumberId);
        document.AddParagraph("Body");

        string html = document.ToHtml(new WordHtmlExportOptions { ExportHeadersAndFooters = true });
        Assert.Contains("<header", html, StringComparison.Ordinal);
        Assert.Contains("<ul", html, StringComparison.Ordinal);
        Assert.Contains("<li>Header list item</li>", html, StringComparison.Ordinal);

        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList pictureList = document.AddPictureBulletList(image, "Kulek.jpg");
        AttachToList(document.Header.Default.AddParagraph("Header picture marker"), pictureList.NumberId);
        var bodyOnly = document.ConvertToHtmlWithEvidence(new WordHtmlExportOptions { ExportHeadersAndFooters = false });
        Assert.DoesNotContain(bodyOnly.Report.Diagnostics, diagnostic => diagnostic.Code == "picture-bullet-text-fallback");
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

        string text = PdfReadDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }), new PdfLoadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Contains(expected, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfHeaderListMarkerPreservesMarkerSpecificFormatting() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "FF0000", global::OfficeIMO.OfficeLength.Points(18));
        list.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordParagraph header = document.Header!.Default!.AddParagraph("HeaderStyledMarker");
        header.FontSize = global::OfficeIMO.OfficeLength.Points(9);
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })), new PdfLoadOptions { IncludeArtifactText = true }).Pages[0].GetTextSpans().ToArray();
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
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(108D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(18D);
        level.LevelSuffix = WordListLevelSuffix.Tab;
        WordParagraph header = document.Header.Default.AddParagraph("IndentedHeaderItem");
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })), new PdfLoadOptions { IncludeArtifactText = true }).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan reference = Assert.Single(spans, span => span.Text.Contains("PlainHeaderReference", StringComparison.Ordinal));
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        PdfTextSpan content = Assert.Single(spans, span => span.Text.Contains("IndentedHeaderItem", StringComparison.Ordinal));
        Assert.True(marker.X > reference.X + 70D, $"reference={reference.X}, marker={marker.X}");
        Assert.InRange(Math.Abs(content.X - (reference.X + 108D)), 0D, 0.5D);
        Assert.True(content.X >= marker.X + marker.Advance);
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
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(36D);
        level.LevelJustification = alignment;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        NumberingSymbolRunProperties markerProperties = level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>() ??
            level.OpenXmlElement.AppendChild(new NumberingSymbolRunProperties());
        markerProperties.Append(new Bold(), new FontSize { Val = "36" });
        WordParagraph nineItem = document.Header!.Default!.AddParagraph("HeaderNine");
        nineItem.FontSize = global::OfficeIMO.OfficeLength.Points(9);
        AttachToList(nineItem, list.NumberId);
        WordParagraph tenItem = document.Header.Default.AddParagraph("HeaderTen");
        tenItem.FontSize = global::OfficeIMO.OfficeLength.Points(9);
        AttachToList(tenItem, list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })), new PdfLoadOptions { IncludeArtifactText = true }).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan nine = Assert.Single(spans, span => span.Text == "9.");
        PdfTextSpan ten = Assert.Single(spans, span => span.Text == "10.");
        PdfTextSpan nineText = Assert.Single(spans, span => span.Text.Contains("HeaderNine", StringComparison.Ordinal));
        PdfTextSpan tenText = Assert.Single(spans, span => span.Text.Contains("HeaderTen", StringComparison.Ordinal));
        double expectedStartShift = alignment == WordListLevelAlignment.Right ? 10D : 5D;
        Assert.InRange(nine.X - ten.X, expectedStartShift - 1D, expectedStartShift + 1D);
        double expectedTextShift = alignment == WordListLevelAlignment.Right ? 0D : expectedStartShift;
        Assert.InRange(tenText.X - nineText.X, expectedTextShift - 1D, expectedTextShift + 1D);
    }

    [Fact]
    public void PdfHeaderExtremeListIndentUsesBoundedMarkerSpacing() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points((int.MaxValue) / 20D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Zero;
        level.LevelSuffix = WordListLevelSuffix.Tab;
        WordParagraph header = document.Header!.Default!.AddParagraph("ExtremeIndentHeader");
        AttachToList(header, list.NumberId);
        document.AddParagraph("Body");

        byte[] pdf = document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        Assert.InRange(pdf.Length, 1, 1_000_000);
        Assert.Contains("ExtremeIndentHeader", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfHeaderTextBoxPreservesInnerMarkerStyle() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('*', "Arial", "FF0000", global::OfficeIMO.OfficeLength.Points(18));
        list.Numbering.Levels[0].OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()?.Append(new Bold());
        WordParagraph host = document.Header!.Default!.AddParagraph("Before ");
        WordTextBox box = host.AddTextBox("StyledHeaderBox", WordImageTextWrapping.Square);
        host.AddText(" After");
        AttachToList(box.Paragraphs[0], list.NumberId);
        document.AddParagraph("Body");

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text == "*");
        Assert.InRange(marker.FontSize, 17.5D, 18.5D);
        Assert.True(marker.IsBold);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), marker.Color);
        string text = PdfReadDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).ExtractText();
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
            PdfReadDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }), new PdfLoadOptions { IncludeArtifactText = true }).Pages[0].GetTextSpans(),
            span => span.Text == "*");
        Assert.True(marker.Advance > 0D);
    }

    [Theory]
    [InlineData(WordListLevelAlignment.Right)]
    [InlineData(WordListLevelAlignment.Center)]
    public void TextBoxNumberingHonorsLevelJustificationInPdf(WordListLevelAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(36D);
        level.LevelJustification = alignment;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        NumberingSymbolRunProperties markerProperties = level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>() ??
            level.OpenXmlElement.AppendChild(new NumberingSymbolRunProperties());
        markerProperties.Append(new Bold(), new FontSize { Val = "36" });
        WordTextBox box = document.AddTextBox("TextBoxNine");
        box.Width = OfficeLength.Centimeters(12D);
        box.RelativeWidthPercentage = 0;
        box.Content!.Append(new Paragraph(new Run(new Text("TextBoxTen"))));
        AttachToList(box.Paragraphs[0], list.NumberId);
        AttachToList(box.Paragraphs[1], list.NumberId);

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan pdfNine = Assert.Single(spans, span => span.Text == "9.");
        PdfTextSpan pdfTen = Assert.Single(spans, span => span.Text == "10.");
        PdfTextSpan pdfNineText = Assert.Single(spans, span => span.Text.Contains("TextBoxNine", StringComparison.Ordinal));
        PdfTextSpan pdfTenText = Assert.Single(spans, span => span.Text.Contains("TextBoxTen", StringComparison.Ordinal));
        double expectedStartShift = alignment == WordListLevelAlignment.Right ? 10D : 5D;
        Assert.InRange(pdfNine.X - pdfTen.X, expectedStartShift - 1D, expectedStartShift + 1D);
        // With no suffix, centered markers end at different body positions;
        // right-aligned markers share their right edge and body position.
        double expectedBodyShift = alignment == WordListLevelAlignment.Center
            ? (pdfTen.Advance - pdfNine.Advance) / 2D : 0D;
        Assert.InRange(Math.Abs(pdfTenText.X - pdfNineText.X - expectedBodyShift), 0D, 0.05D);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelSuffix.Space)]
    [InlineData(WordListLevelSuffix.Tab)]
    public void ImageListMarkersMatchTheSharedPdfPage(WordListLevelSuffix suffix) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "000000");
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        level.IndentationHanging= global::OfficeIMO.OfficeLength.Points(18D);
        level.LevelSuffix = suffix;
        list.AddItem("ImageSuffixBody");
        WordTextBox box = document.AddTextBox("ImageSuffixBox");
        AttachToList(box.Paragraphs[0], list.NumberId);
        WordParagraph cell = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "ImageSuffixCell";
        AttachToList(cell, list.NumberId);

        var layout = new WordPdfExportOptions { IncludePageNumbers = false };
        var options = new WordImageExportOptions { LayoutOptions = layout };
        PdfReadPage pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf(layout).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).Pages);
        WordDocumentVisualInfo snapshot = document.CreateVisualSnapshot(options);
        string text = string.Concat(snapshot.TextSpans.Select(span => span.Text));
        Assert.Contains("ImageSuffixBody", text, StringComparison.Ordinal);
        Assert.Contains("ImageSuffixBox", text, StringComparison.Ordinal);
        Assert.Contains("ImageSuffixCell", text, StringComparison.Ordinal);
        Assert.Contains("*", text, StringComparison.Ordinal);
        Assert.Equal(pdfPage.GetTextSpans().Select(span => (span.Text, span.X, span.Y, span.Advance, span.FontSize)),
            snapshot.TextSpans.Select(span => (span.Text, span.X, span.Y, span.Advance, span.FontSize)));
        Assert.Equal(pdfPage.ExportImage(OfficeImageExportFormat.Svg).Bytes,
            document.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
    }

    [Fact]
    public void ImageMarkerlessListPreservesExplicitZeroIndent() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("PlainZeroReference");
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        list.Numbering.Levels[0].IndentationLeft= global::OfficeIMO.OfficeLength.Zero;
        list.Numbering.Levels[0].IndentationHanging= global::OfficeIMO.OfficeLength.Zero;
        WordParagraph item = document.AddParagraph("MarkerlessZero");
        AttachToList(item, list.NumberId);

        IReadOnlyList<PdfTextSpan> texts = document.CreateVisualSnapshot().TextSpans;
        double referenceX = Assert.Single(texts, text => text.Text == "PlainZeroReference").X;
        double markerlessX = Assert.Single(texts, text => text.Text == "MarkerlessZero").X;
        Assert.InRange(Math.Abs(referenceX - markerlessX), 0D, 0.01D);
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

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("text-box nesting exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PictureBulletFallbackIsReportedByHtmlAndImageExports() {
        using WordDocument document = WordDocument.Create();
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        list.AddItem("Picture item");

        var html = document.ConvertToHtmlWithEvidence();
        Assert.Contains(html.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "picture-bullet-text-fallback" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        WordDocumentVisualInfo snapshot = document.CreateVisualSnapshot();
        Assert.Contains(snapshot.Diagnostics, diagnostic =>
            diagnostic.Code == "native-picture-bullet-text-fallback" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(document.ExportImage(OfficeImageExportFormat.Svg).Diagnostics, diagnostic =>
            diagnostic.Code == "native-picture-bullet-text-fallback");
    }

    [Fact]
    public void TextBoxListIndentationIsKeptPerParagraphInPdfLayout() {
        using WordDocument document = WordDocument.Create();
        WordList bulletList = document.AddCustomBulletList('◆', "Arial", "000000");
        bulletList.Numbering.Levels[0].IndentationLeft= global::OfficeIMO.OfficeLength.Points(90D);
        bulletList.Numbering.Levels[0].IndentationHanging= global::OfficeIMO.OfficeLength.Points(18D);
        WordList markerlessList = document.AddCustomList();
        markerlessList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        markerlessList.Numbering.Levels[0].IndentationLeft= global::OfficeIMO.OfficeLength.Points(108D);
        WordTextBox box = document.AddTextBox("PlainBox");
        box.Width = OfficeLength.Centimeters(12D);
        box.RelativeWidthPercentage = 0;
        box.Content!.Append(new Paragraph(new Run(new Text("BulletBox"))));
        box.Content.Append(new Paragraph(new Run(new Text("UnmarkedBox"))));
        AttachToList(box.Paragraphs[1], bulletList.NumberId);
        AttachToList(box.Paragraphs[2], markerlessList.NumberId);

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].GetTextSpans().ToArray();
        var lines = spans.GroupBy(span => Math.Round(span.Y, 2)).Select(group => new {
            Text = string.Concat(group.OrderBy(span => span.X).Select(span => span.Text)),
            Spans = group.OrderBy(span => span.X).ToArray()
        }).ToArray();
        double WordX(string word) {
            var line = Assert.Single(lines, item => item.Text.Contains(word, StringComparison.Ordinal));
            int offset = line.Text.IndexOf(word, StringComparison.Ordinal), consumed = 0;
            foreach (var span in line.Spans) {
                if (consumed + span.Text.Length > offset) return span.X;
                consumed += span.Text.Length;
            }
            throw new InvalidOperationException("Expected word has no positioned text span.");
        }
        double plainX = WordX("PlainBox"), bulletX = WordX("BulletBox"), unmarkedX = WordX("UnmarkedBox");
        Assert.True(bulletX > plainX + 35D, $"plain={plainX}, bullet={bulletX}");
        Assert.True(unmarkedX > plainX + 55D, $"plain={plainX}, unmarked={unmarkedX}");
    }

    [Fact]
    public void HostListMarkerStaysOutsideItsUnnumberedPdfTextBox() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordParagraph host = document.AddParagraph();
        WordTextBox box = host.AddTextBox("UnnumberedInner", WordImageTextWrapping.Square);
        AttachToList(host, list.NumberId);

        PdfTextSpan[] spans = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text.Contains("1.", StringComparison.Ordinal));
        PdfTextSpan inner = Assert.Single(spans, span => span.Text.Contains("UnnumberedInner", StringComparison.Ordinal));
        Assert.NotEqual(marker.Y, inner.Y);
        Assert.Equal(1, PdfReadDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).ExtractText()
            .Split(new[] { "1." }, StringSplitOptions.None).Length - 1);
        Assert.NotNull(box);
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

        byte[] pdf = document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        (double Start, double End) marker = PdfLetterRun(pdf, "*");
        (double Start, double End) content = PdfLetterRun(pdf, "SuffixCell");
        Assert.InRange(content.Start - marker.End, minimumGap, maximumGap);
    }

    // A line's marker and text can share one PDF show operator, so they are located by their letters.
    private static (double Start, double End) PdfLetterRun(byte[] pdf, string text) {
        using var document = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var letters = document.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToList();
        string joined = string.Concat(letters.Select(letter => letter.Value));
        int index = joined.IndexOf(text, StringComparison.Ordinal);
        Assert.True(index >= 0 && joined.IndexOf(text, index + 1, StringComparison.Ordinal) < 0, "Expected one '" + text + "' in '" + joined + "'.");
        return (letters[index].StartBaseLine.X, letters[index + text.Length - 1].EndBaseLine.X);
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
