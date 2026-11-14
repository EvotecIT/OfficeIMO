using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfUnavailableLinkTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MaskedDrawingLinksRespectSelectedTargetsAndStrictLossHandling(bool nestedClip) {
        var paint = new OfficeDrawing(160D, 40D);
        foreach (var panel in new[] { ("#later", OfficeColor.Red, 0D), ("#first", OfficeColor.Green, 40D),
            ("https://example.test/valid", OfficeColor.Blue, 80D), ("terms.html", OfficeColor.Yellow, 120D) }) {
            OfficeShape shape = OfficeShape.Rectangle(20D, 20D);
            shape.FillColor = panel.Item2;
            paint.AddShape(shape, panel.Item3, 0D);
            paint.AddLink(panel.Item1, panel.Item3, 0D, 20D, 20D);
        }
        if (nestedClip) paint = new OfficeDrawing(160D, 40D)
            .AddClippedDrawing(paint, 0D, 0D, OfficeClipPath.Rectangle(160D, 40D));
        OfficeShape maskShape = OfficeShape.Rectangle(160D, 40D);
        maskShape.FillColor = OfficeColor.White;
        var maskPaint = new OfficeDrawing(160D, 40D).AddShape(maskShape, 0D, 0D);
        var mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(maskPaint,
            OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
        var drawing = new OfficeDrawing(160D, 40D)
            .AddEffectDrawing(paint, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);
        var full = new HtmlRenderDocument(HtmlRenderMode.Paged, new[] {
            new HtmlRenderPage(1, 200D, 100D, new HtmlRenderVisual[] {
                new HtmlRenderNamedDestination("first", 0D, 0D, 0, "first"),
                new HtmlRenderDrawing(drawing, 0D, 0D, 160D, 40D, 1, "Linked panels", null, "masked drawing")
            }),
            new HtmlRenderPage(2, 200D, 100D, new[] { new HtmlRenderNamedDestination("later", 0D, 0D, 0, "later") })
        }, new HtmlDiagnosticReport());
        var options = new HtmlPdfExportOptions();
        HtmlRenderRequest request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options)
            .WithPageSet(HtmlRenderPageSet.Page(0));
        HtmlRenderDocument selected = HtmlRenderRequestProcessor.Complete(request, full, options, CancellationToken.None).Document;

        var result = HtmlPdfRenderedConverter.CreatePdf(selected, options, CancellationToken.None);
        byte[] bytes = result.Document.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(1, info.PageCount);
        Assert.Contains("html-fragment:first", info.LinkDestinationNames);
        Assert.DoesNotContain("html-fragment:later", info.LinkDestinationNames);
        Assert.Contains("https://example.test/valid", info.LinkUris);
        Assert.DoesNotContain("terms.html", info.LinkUris);
        Assert.Contains(result.ConversionReport.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable
            && warning.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(result.ConversionReport.Warnings, warning => warning.Code == "HtmlPdfDrawingEffectRasterized");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 5));

        var strict = new HtmlPdfExportOptions { FidelityPolicy = HtmlRenderFidelityPolicy.FailOnLoss };
        HtmlConversionException failure = Assert.Throws<HtmlConversionException>(() =>
            HtmlPdfRenderedConverter.CreatePdf(selected, strict, CancellationToken.None));
        Assert.Contains(failure.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable);
        var complete = HtmlPdfRenderedConverter.CreatePdf(full, options, CancellationToken.None);
        Assert.Contains("html-fragment:later", PdfInspector.Inspect(complete.Document.ToBytes()).LinkDestinationNames);
    }

    [Theory]
    [InlineData("<a href='terms.html'>Terms</a>")]
    [InlineData("<a href='terms.html'><img alt='Terms' width='20' height='20' src='data:image/png;base64,IMAGE'></a>")]
    [InlineData("<svg width='80' height='40'><a href='terms.html'><rect width='80' height='40' fill='green'/></a></svg>")]
    public void RelativeLinksWithoutABaseDoNotPreventPdfOutput(string body) {
        body = body.Replace("IMAGE", Convert.ToBase64String(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red))));
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p>Body <a href='https://example.test/absolute'>Absolute</a></p>" + body);
        var result = source.ConvertToPdfWithEvidence();

        byte[] bytes = result.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(1, info.PageCount);
        Assert.DoesNotContain("terms.html", info.LinkUris);
        Assert.Contains("https://example.test/absolute", info.LinkUris);
        Assert.Contains("Body", PdfReadDocument.Open(bytes).ExtractText());
        if (body.Contains("<img", StringComparison.Ordinal)) {
            Assert.NotEmpty(PdfReadDocument.Open(bytes).Pages[0].GetImages());
        }
        Assert.Contains(result.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable
            && warning.LossKind == OfficeConversionLossKind.Omission);

        byte[] resolved = source.RenderToPdfBytes(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlPdfExportOptions { BaseUri = new Uri("https://example.test/docs/") }));
        Assert.Contains("https://example.test/docs/terms.html", PdfInspector.Inspect(resolved).LinkUris);
    }

    [Fact]
    public void LinksToUnselectedPagesAreOmittedWhileSelectedTargetsRemainNavigable() {
        const string html = "<style>@page{size:300px 200px;margin:8px}</style>"
            + "<p id='first'><a href='#later'>Later</a> <a href='#first'>First</a></p>"
            + "<h1 id='later' style='break-before:page'>Later page</h1>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderRequest request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf);
        byte[] complete = source.RenderToPdfBytes(request);
        Assert.Equal(2, PdfInspector.Inspect(complete).PageCount);
        Assert.Contains("html-fragment:later", PdfInspector.Inspect(complete).LinkDestinationNames);

        var result = source.RenderToPdfResult(request.WithPageSet(HtmlRenderPageSet.Page(0)));
        byte[] selected = result.Output.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(selected);

        Assert.Equal(1, info.PageCount);
        Assert.Contains("html-fragment:first", info.LinkDestinationNames);
        Assert.DoesNotContain("html-fragment:later", info.LinkDestinationNames);
        Assert.Contains(result.Output.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable
            && warning.LossKind == OfficeConversionLossKind.Omission);
    }
}
