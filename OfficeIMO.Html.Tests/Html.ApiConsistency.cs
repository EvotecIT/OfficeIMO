using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlApiConsistencyTests {
    [Fact]
    public void DirectHtmlOutputs_UseTargetOrDestinationNaming() {
        Type pdfExtensions = typeof(HtmlPdfConverterExtensions);
        Type imageExtensions = typeof(HtmlImageExportExtensions);

        Assert.Contains(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.ToPdf)
            && method.ReturnType == typeof(byte[])
            && method.GetParameters()[0].ParameterType == typeof(string));

        Assert.DoesNotContain(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)
            && method.ReturnType == typeof(byte[]));

        Assert.All(
            pdfExtensions.GetMethods().Where(method => method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)),
            method => Assert.Contains(method.GetParameters()[1].ParameterType, new[] { typeof(string), typeof(Stream) }));

        Type[] sources = { typeof(string), typeof(HtmlConversionDocument), typeof(Stream) };
        foreach (Type source in sources) {
            Assert.Contains(pdfExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
                && method.GetParameters()[0].ParameterType == source
                && method.GetParameters()[1].ParameterType == typeof(string));
            Assert.Contains(pdfExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
                && method.GetParameters()[0].ParameterType == source
                && method.GetParameters()[1].ParameterType == typeof(Stream));

            Assert.Contains(imageExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlImageExportExtensions.ToPngResult)
                && method.ReturnType == typeof(OfficeImageExportResult)
                && method.GetParameters()[0].ParameterType == source);
            Assert.Contains(imageExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlImageExportExtensions.ToSvgResult)
                && method.ReturnType == typeof(OfficeImageExportResult)
                && method.GetParameters()[0].ParameterType == source);
            Assert.Contains(imageExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlImageExportExtensions.ToPngResultsAsync)
                && method.GetParameters()[0].ParameterType == source);
            Assert.Contains(imageExtensions.GetMethods(), method =>
                method.Name == nameof(HtmlImageExportExtensions.ToSvgResultsAsync)
                && method.GetParameters()[0].ParameterType == source);
        }
    }

    [Fact]
    public async Task PdfPngAndSvg_AcceptTheSameHtmlSources() {
        const string html = "<h1>Source parity</h1><p>String, document, and stream.</p>";
        HtmlConversionDocument document = HtmlConversionDocumentBuilder.Build(html);
        var options = new HtmlPdfSaveOptions();

        Assert.NotEmpty(document.ToPdf(options));
        Assert.NotEmpty(document.ToPng(options));
        Assert.StartsWith("<svg", document.ToSvg(options), StringComparison.Ordinal);
        OfficeImageExportResult pngResult = document.ToPngResult(options);
        OfficeImageExportResult svgResult = document.ToSvgResult(options);
        Assert.Equal(OfficeImageExportFormat.Png, pngResult.Format);
        Assert.Equal(OfficeImageExportFormat.Svg, svgResult.Format);
        Assert.NotEmpty(pngResult.Bytes);
        Assert.NotEmpty(svgResult.Bytes);
        Assert.NotNull(pngResult.Diagnostics);
        Assert.NotNull(svgResult.Diagnostics);

        using var pdfSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var pngSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var svgSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var pngResultSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var svgResultSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));

        Assert.NotEmpty(await pdfSource.ToPdfAsync(options));
        Assert.NotEmpty(await pngSource.ToPngAsync(options));
        Assert.StartsWith("<svg", await svgSource.ToSvgAsync(options), StringComparison.Ordinal);
        Assert.Equal(OfficeImageExportFormat.Png, (await pngResultSource.ToPngResultAsync(options)).Format);
        Assert.Equal(OfficeImageExportFormat.Svg, (await svgResultSource.ToSvgResultAsync(options)).Format);
    }

    [Fact]
    public void SharedPdfOptions_PreserveImageModeWhilePdfRemainsPaged() {
        const string html = "<div style='height:1800px;background:#336699'></div>";
        var options = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlPdfSaveOptions clone = options.ClonePdf();
        HtmlPdfSaveOptions copied = new HtmlPdfSaveOptions(options);

        Assert.Equal(HtmlRenderMode.Continuous, clone.Mode);
        Assert.Equal(HtmlRenderMode.Continuous, copied.Mode);
        Assert.Single(HtmlRenderEngine.Render(html, options).Pages);
        Assert.Single(html.ExportImages(OfficeIMO.Drawing.OfficeImageExportFormat.Png, options));
        Assert.Single(html.ExportImages(OfficeIMO.Drawing.OfficeImageExportFormat.Svg, options));
        Assert.True(OfficeIMO.Pdf.PdfInspector.Inspect(html.ToPdf(options)).PageCount > 1);
    }

    [Fact]
    public void HtmlPdfOptions_CopyConstructorPreservesPdfSpecificSettings() {
        var fontFamily = new PdfCore.PdfEmbeddedFontFamily("ContractFont", new byte[] { 1 });
        var shapingProvider = new NullShapingProvider();
        var options = new HtmlPdfSaveOptions {
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            TextShapingMode = PdfCore.PdfTextShapingMode.UnicodeScalar,
            FontFamily = fontFamily,
            TextShapingProvider = shapingProvider
        };

        var copied = new HtmlPdfSaveOptions(options);

        Assert.Equal(PdfCore.PdfTextFallbackFeatures.None, copied.TextFallbacks);
        Assert.Equal(PdfCore.PdfTextShapingMode.UnicodeScalar, copied.TextShapingMode);
        Assert.Same(fontFamily, copied.FontFamily);
        Assert.Same(shapingProvider, copied.TextShapingProvider);
    }

    [Fact]
    public void HtmlPdf_ByteSerializationHonorsCancellation() {
        PdfCore.PdfDocumentConversionResult result = "<p>Cancellation contract</p>".ToPdfResult();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => {
            _ = HtmlPdfConverterExtensions.SerializeToBytes(result, cancellation.Token);
        });
    }

    [Fact]
    public void RawTextConverters_UseSourceExplicitPdfNames() {
        Assert.DoesNotContain(typeof(MarkdownPdfConverterExtensions).GetMethods(), method =>
            method.Name == "ToPdf" && method.GetParameters()[0].ParameterType == typeof(string));
        Assert.Contains(typeof(MarkdownPdfConverterExtensions).GetMethods(), method =>
            method.Name == "ToPdfFromMarkdown" && method.GetParameters()[0].ParameterType == typeof(string));

    }

    [Fact]
    public void HtmlPdf_DependsOnlyOnDirectRenderingOwners() {
        string[] references = typeof(HtmlPdfSaveOptions).Assembly
            .GetReferencedAssemblies()
            .Select(reference => reference.Name ?? string.Empty)
            .ToArray();

        Assert.Contains("OfficeIMO.Html", references);
        Assert.Contains("OfficeIMO.Drawing", references);
        Assert.Contains("OfficeIMO.Pdf", references);
        Assert.DoesNotContain("OfficeIMO.Markdown.Html", references);
        Assert.DoesNotContain("OfficeIMO.Markdown.Pdf", references);
        Assert.DoesNotContain("OfficeIMO.Word.Html", references);
        Assert.DoesNotContain("OfficeIMO.Word.Pdf", references);
    }

    private sealed class NullShapingProvider : PdfCore.IPdfTextShapingProvider {
        public PdfCore.PdfTextShapingResult? ShapeText(PdfCore.PdfTextShapingRequest request) => null;
    }
}
