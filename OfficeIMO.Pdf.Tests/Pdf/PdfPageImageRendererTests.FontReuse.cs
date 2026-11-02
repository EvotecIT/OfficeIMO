using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedFontAliasesKeepTextAndPaintedGlyphsAcrossDiagnosticOrderAndCancellation(bool diagnosticsFirst) {
        byte[] pdf = BuildFontAliasPdf(embedded: true);
        PdfDocument document = PdfDocument.Load(pdf);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Render.Pages(cancellationToken: cancellation.Token));

        string[]? diagnostics = diagnosticsFirst ? FontAliasDiagnostics(document) : null;
        OfficeDrawing drawing = document.Render.Drawing(1);
        diagnostics ??= FontAliasDiagnostics(document);
        Assert.Empty(diagnostics);
        Assert.Equal("ffiffiR", string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text)));

        byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing, background: OfficeColor.White);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeDrawing freshDrawing = PdfDocument.Load(pdf).Render.Drawing(1);
        Assert.Equal(svg, OfficeDrawingSvgExporter.ToSvg(freshDrawing));
        Assert.Equal(png, OfficeDrawingRasterRenderer.ToPng(freshDrawing, background: OfficeColor.White));
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        // A and B have deliberately different outlines. Both aliases of A stay narrow, while
        // the distinct font dictionary named F1 in the second form paints the wider B glyph.
        Assert.Equal(OfficeColor.Black, raster!.GetPixel(23, 90));
        Assert.Equal(OfficeColor.White, raster.GetPixel(29, 90));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(83, 90));
        Assert.Equal(OfficeColor.White, raster.GetPixel(89, 90));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(149, 90));

        Assert.Throws<OperationCanceledException>(() => document.Render.Pages(cancellationToken: cancellation.Token));
        for (int repeat = 0; repeat < 2; repeat++) {
            Assert.Equal(diagnostics, FontAliasDiagnostics(document));
            OfficeDrawing repeated = document.Render.Drawing(1);
            Assert.Equal(svg, OfficeDrawingSvgExporter.ToSvg(repeated));
            Assert.Equal(png, OfficeDrawingRasterRenderer.ToPng(repeated, background: OfficeColor.White));
        }

        // The public batch path requests diagnostics before drawing. Its output must also survive
        // cold and warm cancellation, independently of the drawing-first export contract.
        PdfPageRenderResult batch = Assert.Single(document.Render.Pages(options: new PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png }));
        Assert.True(batch.Succeeded, string.Join(" ", batch.Diagnostics));
        PdfPageRenderResult fresh = Assert.Single(PdfDocument.Load(pdf).Render.Pages(options: new PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png }));
        Assert.Equal(fresh.Bytes, batch.Bytes);
        Assert.Equal(fresh.CapabilityDiagnostics.Select(DiagnosticContract), batch.CapabilityDiagnostics.Select(DiagnosticContract));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedUnembeddedFontReportsEachInvokedLocalAliasAcrossDiagnosticOrder(bool diagnosticsFirst) {
        PdfDocument document = PdfDocument.Load(BuildFontAliasPdf(embedded: false));
        string[]? diagnostics = diagnosticsFirst ? FontAliasDiagnostics(document) : null;
        OfficeDrawing drawing = document.Render.Drawing(1);
        diagnostics ??= FontAliasDiagnostics(document);

        Assert.Equal(new[] { "render.resource.font-substitution|F1", "render.resource.font-substitution|Nested" },
            diagnostics.Where(value => value.StartsWith("render.resource.font-substitution|", StringComparison.Ordinal)).OrderBy(value => value, StringComparer.Ordinal));
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Equal(diagnostics, FontAliasDiagnostics(document));
        Assert.Equal(svg, OfficeDrawingSvgExporter.ToSvg(document.Render.Drawing(1)));
    }

    private static string[] FontAliasDiagnostics(PdfDocument document) =>
        document.Render.CapabilityDiagnostics(1).Select(DiagnosticContract).ToArray();

    private static string DiagnosticContract(PdfRenderCapabilityDiagnostic diagnostic) =>
        diagnostic.Code + "|" + diagnostic.Subject;

    private static byte[] BuildFontAliasPdf(bool embedded) {
        string firstFont = embedded
            ? "<< /Type /Font /Subtype /TrueType /BaseFont /SharedBase /FirstChar 65 /LastChar 65 /Widths [500] /FontDescriptor 7 0 R /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /A] >> /ToUnicode 9 0 R >>"
            : "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>";
        string secondFont = embedded
            ? "<< /Type /Font /Subtype /TrueType /BaseFont /SharedBase /FirstChar 65 /LastChar 65 /Widths [500] /FontDescriptor 7 0 R /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /B] >> /ToUnicode 10 0 R >>"
            : firstFont;
        string descriptor = "7 0 obj\n<< /Type /FontDescriptor /FontName /SharedBase /Flags 32 /FontBBox [0 0 600 700] /Ascent 800 /Descent -200 /CapHeight 700 /ItalicAngle 0 /StemV 80 /FontFile2 8 0 R >>\nendobj";
        string program = BuildStreamObject(8, "<< /Filter /ASCIIHexDecode", BitConverter.ToString(
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B')).Replace("-", string.Empty) + ">");
        const string cmapPrefix = "begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfchar\n<41> <";
        const string cmapSuffix = ">\nendbfchar\nendcmap";
        string sharedForm = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << /Font << /Nested 5 0 R >> >>",
            "BT /Nested 20 Tf 1 0 0 1 80 100 Tm (A) Tj ET");
        string otherForm = BuildStreamObject(12, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << /Font << /F1 6 0 R >> >>",
            "BT /F1 20 Tf 1 0 0 1 140 100 Tm (A) Tj ET");
        return BuildSingleStreamPdf("BT /F1 20 Tf 1 0 0 1 20 100 Tm (A) Tj ET /SharedForm Do /OtherForm Do",
            "<< /Font << /F1 5 0 R >> /XObject << /SharedForm 11 0 R /OtherForm 12 0 R >> >>",
            "5 0 obj\n" + firstFont + "\nendobj", "6 0 obj\n" + secondFont + "\nendobj", descriptor, program,
            BuildStreamObject(9, "<<", cmapPrefix + "006600660069" + cmapSuffix),
            BuildStreamObject(10, "<<", cmapPrefix + "0052" + cmapSuffix), sharedForm, otherForm);
    }
}
