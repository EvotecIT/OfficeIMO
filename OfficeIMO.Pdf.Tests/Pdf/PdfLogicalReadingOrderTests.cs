using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;
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

    private static void AssertInOrder(string value, params string[] markers) {
        int previous = -1;
        foreach (string marker in markers) {
            int current = value.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(current > previous, "Expected marker '" + marker + "' after the previous marker.");
            previous = current;
        }
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
}
