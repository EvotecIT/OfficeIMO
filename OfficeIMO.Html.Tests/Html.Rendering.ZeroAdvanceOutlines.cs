using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlShaping_ZeroAdvanceOutlinesRetainInkAndSourceAcrossBodyAndSvg(bool svg) {
        const string mark = "\u0301";
        string content = svg
            ? "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='50' viewBox='0 0 80 50'>"
                + "<text x='20' y='30' font-family='Typography' font-size='18'>" + mark + "</text></svg>"
            : mark;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", string.Empty, "width:auto", content));
        HtmlToPdfOptions options = PageMarginShapingOptions(cff: true);
        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        byte[] bytes = pdf.ToBytes();
        pdf.RequireNoLoss();

        Assert.Equal(mark, PdfCore.PdfReadDocument.Open(bytes).ExtractText().Trim());
        OfficeDrawing drawing = PdfCore.PdfDocument.Load(bytes).Render.Drawing(1);
        Assert.Contains(DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingShape>(), shape =>
            shape.Shape.PathCommands.Count > 0 && shape.Shape.FillColor is OfficeColor fill
                && fill.A > 0 && fill.R < 128 && fill.G < 128 && fill.B < 128);

        OfficeImageExportResult image = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(pdf,
            OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, image.Bytes);
        Assert.DoesNotContain(image.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
        Assert.DoesNotContain(replay.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
    }
}
