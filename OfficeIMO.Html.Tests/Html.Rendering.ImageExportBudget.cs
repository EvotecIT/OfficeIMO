using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png, HtmlRenderEncoder.Png)]
    [InlineData(OfficeImageExportFormat.Svg, HtmlRenderEncoder.Svg)]
    public async Task HtmlImages_AggregateEncodedByteLimitAppliesToLegacyAndExplicitRequests(
        OfficeImageExportFormat format,
        HtmlRenderEncoder encoder) {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<html style='margin:0'><body style='margin:0'><img src='data:image/png;base64," + image
            + "' style='display:block;width:40px;height:90px'></body></html>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch,
                40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            MaximumDegreeOfParallelism = 2
        };

        IReadOnlyList<OfficeImageExportResult> baseline = document.ExportImages(format, options);
        Assert.Equal(3, baseline.Count);
        long limit = baseline.Max(result => result.EncodedLength) + 1L;
        Assert.True(baseline.Sum(result => result.EncodedLength) > limit);
        options.MaximumTotalEncodedBytes = limit;

        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, encoder, options).WithPageSet(HtmlRenderPageSet.All());
        OfficeImageExportBatchLimitException legacy = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
            document.ExportImages(format, options));
        OfficeImageExportBatchLimitException explicitRequest = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
            document.RenderImages(request));
        OfficeImageExportBatchLimitException retainedResult = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
            HtmlRenderEngine.Execute(document, request).ExportImages());
        OfficeImageExportBatchLimitException asynchronous = await Assert.ThrowsAsync<OfficeImageExportBatchLimitException>(() =>
            document.ExportImagesAsync(format, options));

        Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), legacy.LimitName);
        Assert.Equal(legacy.LimitName, explicitRequest.LimitName);
        Assert.Equal(legacy.LimitName, retainedResult.LimitName);
        Assert.Equal(legacy.LimitName, asynchronous.LimitName);
    }
    [Fact]
    public async Task HtmlImages_RasterPixelLimitAppliesToIntermediateEffectSurfaces() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>body{margin:0;width:100px;height:100px}"
            + ".layer{position:absolute;left:0;top:0;width:100px;height:100px;opacity:.5;background:red}</style>"
            + "<body><div class='layer'></div><div class='layer'></div></body>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            MaximumRasterPixels = 50_000L
        };
        OfficeImageExportResult baseline = document.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.Equal(100, baseline.Width);
        Assert.Equal(100, baseline.Height);

        options.MaximumRasterPixels = 15_000L;
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options);
        OfficeImageExportLimitException legacy = Assert.Throws<OfficeImageExportLimitException>(() =>
            document.ExportImage(OfficeImageExportFormat.Png, options));
        OfficeImageExportLimitException explicitRequest = Assert.Throws<OfficeImageExportLimitException>(() =>
            document.RenderImages(request));
        OfficeImageExportLimitException retainedResult = Assert.Throws<OfficeImageExportLimitException>(() =>
            HtmlRenderEngine.Execute(document, request).ExportImages());
        OfficeImageExportLimitException asynchronous = await Assert.ThrowsAsync<OfficeImageExportLimitException>(() =>
            document.ExportImageAsync(OfficeImageExportFormat.Png, options));

        foreach (OfficeImageExportLimitException failure in new[] { legacy, explicitRequest, retainedResult, asynchronous }) {
            // The final page has only 10,000 pixels. All public routes must also
            // reject the larger intermediate effect work in the PDF page owner.
            Assert.True(failure.RequestedPixels > 15_000L);
            Assert.Equal(legacy.RequestedPixels, failure.RequestedPixels);
            Assert.Equal(15_000L, failure.MaximumPixels);
        }
    }

    [Fact]
    public async Task HtmlImages_SourcePixelsAreBoundedBeforePdfImagePreparation() {
        string source = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(40, 40));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<body style='margin:0'><img src='data:image/png;base64," + source +
            "' style='display:block;width:10px;height:10px'></body>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 10D,
            ViewportHeight = 10D,
            Margins = HtmlRenderMargins.All(0D),
            MaximumRasterPixels = 1600L,
            Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        };

        var pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 10D,
            ViewportHeight = 10D,
            Margins = HtmlRenderMargins.All(0D),
            MaximumRasterPixels = 1600L
        };
        Assert.DoesNotContain(document.ToPdfDocumentResult(pdfOptions).Report.Warnings,
            warning => warning.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted);
        pdfOptions.MaximumRasterPixels = 1599L;
        Assert.Contains(document.ToPdfDocumentResult(pdfOptions).Report.Warnings,
            warning => warning.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted);
        options.MaximumRasterPixels = 1599L;
        OfficeImageExportPolicyException synchronous = Assert.Throws<OfficeImageExportPolicyException>(() =>
            document.ExportImage(OfficeImageExportFormat.Png, options));
        OfficeImageExportPolicyException asynchronous = await Assert.ThrowsAsync<OfficeImageExportPolicyException>(() =>
            document.ExportImageAsync(OfficeImageExportFormat.Png, options));
        foreach (OfficeImageExportPolicyException failure in new[] { synchronous, asynchronous }) {
            Assert.Contains(failure.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
    }

}
