using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalReadingOrderTests {
    [Fact]
    public void Analyze_OrdersColumnContentBeforeMovingToTheNextColumn() {
        byte[] pdf = BuildPositionedTextPdf(
            "BT /F1 11 Tf\n" +
            "1 0 0 1 30 270 Tm (Left top) Tj\n" +
            "1 0 0 1 230 240 Tm (Right top) Tj\n" +
            "1 0 0 1 30 180 Tm (Left bottom) Tj\n" +
            "1 0 0 1 230 150 Tm (Right bottom) Tj ET");
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);

        PdfLogicalReadingOrderItem[] ordered = PdfLogicalReadingOrderAnalysis.Analyze(page)
            .Where(static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem)
            .ToArray();
        string[] text = ordered.Select(item => GetText(page, item)).ToArray();

        Assert.Equal(new[] { "Left top", "Left bottom", "Right top", "Right bottom" }, text);
        Assert.Equal(new[] { 0, 0, 1, 1 }, ordered.Select(static item => item.ColumnIndex));
        Assert.All(ordered, item => {
            Assert.True(item.HasGeometry);
            Assert.InRange(item.Confidence, 0D, 1D);
            Assert.Contains(item.Evidence, static evidence => evidence.Code == "reading-order.column-band");
        });
    }

    [Fact]
    public void Analyze_KeepsIndentedBodyInItsColumnAndTreatsCenteredTitleAsABandDivider() {
        byte[] pdf = BuildPositionedTextPdf(
            "BT /F1 11 Tf\n" +
            "1 0 0 1 160 300 Tm (Centered title) Tj\n" +
            "1 0 0 1 30 270 Tm (Left top) Tj\n" +
            "1 0 0 1 230 240 Tm (Right top) Tj\n" +
            "1 0 0 1 60 210 Tm (Indented left) Tj\n" +
            "1 0 0 1 30 180 Tm (Left bottom) Tj\n" +
            "1 0 0 1 230 150 Tm (Right bottom) Tj ET");
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);

        PdfLogicalReadingOrderItem[] ordered = PdfLogicalReadingOrderAnalysis.Analyze(page)
            .Where(static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem)
            .ToArray();

        Assert.Equal(
            new[] { "Centered title", "Left top", "Indented left", "Left bottom", "Right top", "Right bottom" },
            ordered.Select(item => GetText(page, item)));
        Assert.True(ordered[0].SpansColumns);
        Assert.Equal(new[] { 0, 0, 0, 0, 1, 1 }, ordered.Select(static item => item.ColumnIndex));
    }

    [Fact]
    public void Analyze_CollapsesRepeatedIndentsAndKeepsANarrowCenteredHeadingOutsideColumns() {
        byte[] pdf = BuildPositionedTextPdf(
            "BT /F1 18 Tf\n" +
            "1 0 0 1 205 580 Tm (Title) Tj\n" +
            "/F1 11 Tf\n" +
            "1 0 0 1 30 520 Tm (Left top) Tj\n" +
            "1 0 0 1 230 500 Tm (Right top) Tj\n" +
            "1 0 0 1 60 430 Tm (Indented one) Tj\n" +
            "1 0 0 1 230 385 Tm (Right middle) Tj\n" +
            "1 0 0 1 60 340 Tm (Indented two) Tj\n" +
            "1 0 0 1 30 250 Tm (Left bottom) Tj\n" +
            "1 0 0 1 230 220 Tm (Right bottom) Tj ET",
            height: 620);
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);

        PdfLogicalReadingOrderItem[] ordered = PdfLogicalReadingOrderAnalysis.Analyze(page)
            .Where(static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem)
            .ToArray();

        Assert.Equal(
            new[] { "Title", "Left top", "Indented one", "Indented two", "Left bottom", "Right top", "Right middle", "Right bottom" },
            ordered.Select(item => GetText(page, item)));
        Assert.True(ordered[0].SpansColumns);
        Assert.Equal(new[] { 0, 0, 0, 0, 0, 1, 1, 1 }, ordered.Select(static item => item.ColumnIndex));
    }

    private static string GetText(PdfLogicalPage page, PdfLogicalReadingOrderItem item) => item.Kind switch {
        PdfLogicalReadingOrderKind.TextBlock => page.TextBlocks[item.SourceIndex].Text,
        PdfLogicalReadingOrderKind.Heading => page.Headings[item.SourceIndex].Line.Text,
        PdfLogicalReadingOrderKind.Paragraph => string.Join(" ", page.Paragraphs[item.SourceIndex].Lines.Select(static line => line.Text)),
        PdfLogicalReadingOrderKind.ListItem => page.ListItems[item.SourceIndex].Line.Text,
        _ => string.Empty
    };

    [Fact]
    public void Analyze_NormalizesRotationAndReportsCropClipping() {
        byte[] source = BuildPositionedTextPdf(
            "BT /F1 12 Tf 12 160 Td (Partially clipped reading-order marker) Tj ET",
            width: 300,
            height: 240);
        byte[] croppedAndRotated = PdfDocument.Load(source)
            .Pages.SetCropBox(40, 0, 280, 240)
            .Pages.Rotate(90, "1")
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(croppedAndRotated).Pages);

        PdfLogicalReadingOrderItem item = Assert.Single(
            PdfLogicalReadingOrderAnalysis.Analyze(page),
            static candidate => candidate.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem);

        Assert.True(item.HasGeometry);
        Assert.True(item.IsClipped);
        Assert.True(item.Left >= 0D && item.Top >= 0D);
        Assert.Contains(item.Evidence, static evidence => evidence.Code == "reading-order.page-rotation");
        Assert.Contains(item.Evidence, static evidence => evidence.Code == "reading-order.crop-clipped");
    }

    [Fact]
    public void Analyze_RetainsVisualOrderWhenRotationDiffersFromRawCanonicalOrder() {
        byte[] source = BuildPositionedTextPdf(
            "BT /F1 12 Tf\n" +
            "1 0 0 1 30 180 Tm (Raw left) Tj\n" +
            "1 0 0 1 230 180 Tm (Raw right) Tj ET");
        byte[] rotated = PdfDocument.Load(source).Pages.Rotate(90, "1").ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(rotated);
        PdfLogicalPage page = Assert.Single(logical.Pages);

        PdfLogicalReadingOrderItem[] ordered = PdfLogicalReadingOrderAnalysis.Analyze(page)
            .Where(static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem)
            .ToArray();

        Assert.Equal(new[] { "Raw right", "Raw left" }, ordered.Select(item => GetText(page, item)));
        AssertInOrder(logical.Text, "Raw right", "Raw left");
    }

    [Fact]
    public void WordAndHtml_ConsumeTheSameColumnReadingOrder() {
        byte[] pdf = BuildPositionedTextPdf(
            "BT /F1 11 Tf\n" +
            "1 0 0 1 30 270 Tm (Left top) Tj\n" +
            "1 0 0 1 230 240 Tm (Right top) Tj\n" +
            "1 0 0 1 30 180 Tm (Left bottom) Tj\n" +
            "1 0 0 1 230 150 Tm (Right bottom) Tj ET");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

        using (OfficeIMO.Word.WordDocument word = logical.ToWordDocument(new PdfWordImportOptions { UseSharedPageReadingOrder = true })) {
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
            string wordText = string.Join(" ", package.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(static text => text.Text));
            AssertInOrder(wordText, "Left top", "Left bottom", "Right top", "Right bottom");
        }

        string html = logical.ToHtml(new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            UseSharedPageReadingOrder = true
        });
        AssertInOrder(html, "Left top", "Left bottom", "Right top", "Right bottom");
    }

    [Fact]
    public void ArtifactConverters_UsePageContentOrderForClassifiedHeadersAndFooters() {
        byte[] pdf = PdfDocument.Create()
            .Header(header => header.AlignLeft().Text("Artifact header {page}/{pages}"))
            .Footer(footer => footer.AlignLeft().Text("Artifact footer {page}/{pages}"))
            .Paragraph(paragraph => paragraph.Text("First page body marker."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page body marker."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Third page body marker."))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);
        PdfLogicalPage firstPage = logical.Pages[0];
        string header = Assert.Single(firstPage.Headers).Text;
        string footer = Assert.Single(firstPage.Footers).Text;

        Assert.DoesNotContain(
            PdfLogicalReadingOrderAnalysis.Analyze(firstPage),
            item => item.Kind == PdfLogicalReadingOrderKind.TextBlock &&
                firstPage.TextBlocks[item.SourceIndex].Kind is PdfLogicalElementKind.Header or PdfLogicalElementKind.Footer);
        Assert.Contains(
            PdfLogicalReadingOrderAnalysis.Analyze(firstPage, PdfLogicalReadingOrderScope.PageContent),
            item => item.Kind == PdfLogicalReadingOrderKind.TextBlock &&
                firstPage.TextBlocks[item.SourceIndex].Kind == PdfLogicalElementKind.Header);

        using (OfficeIMO.Word.WordDocument word = logical.ToWordDocument(new PdfWordImportOptions { UseSharedPageReadingOrder = true })) {
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
            string wordText = string.Join(" ", package.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(static text => text.Text));
            AssertArtifactSequence(wordText, header, "First page body marker.", footer);
        }

        string html = logical.ToHtml(new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            UseSharedPageReadingOrder = true
        });
        AssertArtifactSequence(html, header, "First page body marker.", footer);
        Assert.Contains("<header class=\"pdf-header\">" + header + "</header>", html, StringComparison.Ordinal);
        Assert.Contains("<footer class=\"pdf-footer\">" + footer + "</footer>", html, StringComparison.Ordinal);

        byte[] odt = logical.ToOdtDocument().ToBytes();
        AssertArtifactSequence(ReadOpenDocumentText(odt), header, "First page body marker.", footer);
    }

    [Fact]
    public void CanonicalText_PreservesBodyTextOverAFullPageImage() {
        byte[] pdf = PdfDocument.Create()
            .BackgroundImage(PdfPngTestImages.CreateRgbPng(1, 1), OfficeIMO.Drawing.OfficeImageFit.Stretch, opacity: 0.2)
            .H1("Readable heading")
            .Paragraph(paragraph => paragraph.Text("Readable body text."))
            .ToBytes();

        PdfDocumentReadResult logical = PdfDocument.Load(pdf).Read();

        Assert.Contains("Readable heading", logical.Text, StringComparison.Ordinal);
        Assert.Contains("Readable body text.", logical.Text, StringComparison.Ordinal);
        Assert.Contains(
            PdfLogicalReadingOrderAnalysis.Analyze(Assert.Single(logical.Pages), PdfLogicalReadingOrderScope.PageContent),
            static item => item.Kind == PdfLogicalReadingOrderKind.Image);
    }

    [Fact]
    public void Analyze_TreatsAFullWidthInlineImageAsAColumnBandDivider() {
        byte[] pdf = BuildPositionedTextAndImagePdf(
            "BT /F1 11 Tf\n" +
            "1 0 0 1 30 350 Tm (Upper left) Tj\n" +
            "1 0 0 1 230 340 Tm (Upper right) Tj\n" +
            "1 0 0 1 30 110 Tm (Lower left) Tj\n" +
            "1 0 0 1 230 100 Tm (Lower right) Tj ET\n" +
            "q 380 0 0 60 20 170 cm /Im1 Do Q");
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);

        PdfLogicalReadingOrderItem[] ordered = PdfLogicalReadingOrderAnalysis.Analyze(page, PdfLogicalReadingOrderScope.PageContent)
            .Where(static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.ListItem or PdfLogicalReadingOrderKind.Image)
            .ToArray();
        string[] labels = ordered.Select(item => item.Kind == PdfLogicalReadingOrderKind.Image ? "[image]" : GetText(page, item)).ToArray();

        Assert.Equal(new[] { "Upper left", "Upper right", "[image]", "Lower left", "Lower right" }, labels);
        Assert.True(Assert.Single(ordered, static item => item.Kind == PdfLogicalReadingOrderKind.Image).SpansColumns);
    }

    [Fact]
    public void CanonicalText_PreservesBlocksOmittedByATableProjection() {
        var style = new PdfTableStyle {
            HeaderRowCount = 1,
            CellFills = new Dictionary<(int Row, int Column), PdfColor> {
                [(1, 0)] = new PdfColor(0.86D, 0.92D, 0.99D),
                [(3, 0)] = new PdfColor(0.99D, 0.95D, 0.78D)
            },
            CellAlignments = new Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(1, 0)] = PdfColumnAlign.Center,
                [(3, 0)] = PdfColumnAlign.Center
            },
            CellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(1, 0)] = PdfCellVerticalAlign.Middle
            }
        };
        var rows = new[] {
            new[] { PdfTableCell.TextCell("Service"), PdfTableCell.TextCell("Status"), PdfTableCell.TextCell("Owner") },
            new[] { PdfTableCell.Span(new[] { new PdfTextRun("Identity systems", bold: true, italic: true, fontSize: 17) }, 3) },
            new[] { PdfTableCell.TextCell("Entra"), PdfTableCell.TextCell("Watch"), PdfTableCell.TextCell("IAM") },
            new[] { PdfTableCell.Span(new[] { new PdfTextRun("Follow-up", bold: true, fontSize: 15) }, 3) },
            new[] { PdfTableCell.TextCell("Release"), PdfTableCell.TextCell("Ready"), PdfTableCell.TextCell("OfficeIMO") }
        };
        byte[] pdf = PdfDocument.Create()
            .Table(rows, style: style)
            .Paragraph(paragraph => paragraph.Text("Content below the table"))
            .ToBytes();

        PdfDocumentReadResult logical = PdfDocument.Load(pdf).Read();

        Assert.Contains(logical.TextBlocks, static block => block.Text == "Identity systems");
        Assert.Contains("Identity systems", logical.Text, StringComparison.Ordinal);
        Assert.Contains("Follow-up", logical.Text, StringComparison.Ordinal);
        AssertInOrder(logical.Text, "Identity systems", "Content below the table");
        AssertInOrder(logical.Text, "Follow-up", "Content below the table");
    }

    private static void AssertInOrder(string value, params string[] markers) {
        int previous = -1;
        foreach (string marker in markers) {
            int current = value.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(current > previous, "Expected marker '" + marker + "' after the previous marker.");
            previous = current;
        }
    }

    private static void AssertArtifactSequence(string value, string header, string body, string footer) {
        AssertInOrder(value, header, body, footer);
        Assert.Equal(1, CountOccurrences(value, header));
        Assert.Equal(1, CountOccurrences(value, body));
        Assert.Equal(1, CountOccurrences(value, footer));
    }

    private static int CountOccurrences(string value, string marker) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += marker.Length;
        }
        return count;
    }

    private static string ReadOpenDocumentText(byte[] artifact) {
        using var archive = new ZipArchive(new MemoryStream(artifact), ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry content = archive.GetEntry("content.xml") ?? throw new InvalidDataException("OpenDocument package did not contain content.xml.");
        using var reader = new StreamReader(content.Open());
        return Regex.Replace(reader.ReadToEnd(), "<[^>]+>", " ");
    }

    private static byte[] BuildPositionedTextPdf(string content, int width = 420, int height = 320) {
        int length = Encoding.ASCII.GetByteCount(content);
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + width + " " + height + "] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static byte[] BuildPositionedTextAndImagePdf(string content) => Encoding.ASCII.GetBytes(string.Join("\n", new[] {
        "%PDF-1.4",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 400] /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >> /Contents 4 0 R >>", "endobj",
        "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(content) + " >>", "stream", content, "endstream", "endobj",
        "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
        "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "RGB", "endstream", "endobj",
        "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
    }));
}
