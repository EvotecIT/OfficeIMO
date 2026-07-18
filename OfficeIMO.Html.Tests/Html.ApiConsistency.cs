using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlApiConsistencyTests {
    [Fact]
    public void DirectHtmlOutputs_RequireTheNativeHtmlSourceModel() {
        Type pdfExtensions = typeof(HtmlPdfConverterExtensions);
        Type imageExtensions = typeof(HtmlImageExportExtensions);

        Assert.Contains(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.ToPdf)
            && method.ReturnType == typeof(byte[])
            && method.GetParameters()[0].ParameterType == typeof(HtmlConversionDocument));

        Assert.DoesNotContain(pdfExtensions.GetMethods(), method =>
            method.GetParameters().Length > 0
            && (method.GetParameters()[0].ParameterType == typeof(string)
                || method.GetParameters()[0].ParameterType == typeof(Stream)));

        Assert.DoesNotContain(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)
            && method.ReturnType == typeof(byte[]));

        Assert.All(
            pdfExtensions.GetMethods().Where(method => method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)),
            method => Assert.Contains(method.GetParameters()[1].ParameterType, new[] { typeof(string), typeof(Stream) }));

        Assert.Contains(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
            && method.GetParameters()[0].ParameterType == typeof(HtmlConversionDocument)
            && method.GetParameters()[1].ParameterType == typeof(string));
        Assert.Contains(pdfExtensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
            && method.GetParameters()[0].ParameterType == typeof(HtmlConversionDocument)
            && method.GetParameters()[1].ParameterType == typeof(Stream));

        MethodInfo[] imageMethods = imageExtensions.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);
        Assert.All(
            imageMethods.Where(method => method.GetParameters().Length > 0),
            method => Assert.Equal(typeof(HtmlConversionDocument), method.GetParameters()[0].ParameterType));
        Assert.DoesNotContain(imageMethods, method =>
            method.Name.Contains("Result", StringComparison.Ordinal));
        Assert.All(
            imageMethods.Where(method => method.Name is "SaveAsPng" or "SaveAsSvg"),
            method => Assert.Equal(typeof(OfficeImageExportResult), method.ReturnType));
        Assert.All(
            imageMethods.Where(method => method.Name is "SaveAsPngAsync" or "SaveAsSvgAsync"),
            method => Assert.Equal(typeof(Task<OfficeImageExportResult>), method.ReturnType));
    }

    [Fact]
    public async Task PdfPngAndSvg_AcceptTheSameHtmlSources() {
        const string html = "<h1>Source parity</h1><p>String, document, and stream.</p>";
        HtmlConversionDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html);
        var options = new HtmlPdfSaveOptions();

        Assert.NotEmpty(document.ToPdf(options));
        Assert.NotEmpty(document.ToPng(options));
        Assert.StartsWith("<svg", document.ToSvg(options), StringComparison.Ordinal);
        OfficeImageExportResult pngResult = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult svgResult = document.ExportImage(OfficeImageExportFormat.Svg, options);
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

        HtmlConversionDocument pdfDocument = await HtmlConversionDocument.LoadAsync(pdfSource);
        HtmlConversionDocument pngDocument = await HtmlConversionDocument.LoadAsync(pngSource);
        HtmlConversionDocument svgDocument = await HtmlConversionDocument.LoadAsync(svgSource);
        HtmlConversionDocument pngResultDocument = await HtmlConversionDocument.LoadAsync(pngResultSource);
        HtmlConversionDocument svgResultDocument = await HtmlConversionDocument.LoadAsync(svgResultSource);
        Assert.NotEmpty(await pdfDocument.ToPdfAsync(options));
        Assert.NotEmpty(await pngDocument.ToPngAsync(options));
        Assert.StartsWith("<svg", await svgDocument.ToSvgAsync(options), StringComparison.Ordinal);
        Assert.Equal(OfficeImageExportFormat.Png, (await pngResultDocument.ExportImageAsync(OfficeImageExportFormat.Png, options)).Format);
        Assert.Equal(OfficeImageExportFormat.Svg, (await svgResultDocument.ExportImageAsync(OfficeImageExportFormat.Svg, options)).Format);
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
        Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options).Pages);
        Assert.Single(HtmlConversionDocument.Parse(html).ExportImages(OfficeIMO.Drawing.OfficeImageExportFormat.Png, options));
        Assert.Single(HtmlConversionDocument.Parse(html).ExportImages(OfficeIMO.Drawing.OfficeImageExportFormat.Svg, options));
        Assert.True(OfficeIMO.Pdf.PdfInspector.Inspect(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(options)).PageCount > 1);
    }

    [Fact]
    public void HtmlPdfOptions_CopyConstructorPreservesPdfSpecificSettings() {
        var fontFamily = new PdfCore.PdfEmbeddedFontFamily("ContractFont", new byte[] { 1 });
        var shapingProvider = new NullShapingProvider();
        var options = new HtmlPdfSaveOptions {
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            TextShapingMode = PdfCore.PdfTextShapingMode.UnicodeScalar,
            FontFamily = fontFamily,
            TextShapingProvider = shapingProvider,
            UrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile()
        };

        var copied = new HtmlPdfSaveOptions(options);

        Assert.Equal(PdfCore.PdfTextFallbackFeatures.None, copied.TextFallbacks);
        Assert.Equal(PdfCore.PdfTextShapingMode.UnicodeScalar, copied.TextShapingMode);
        Assert.Same(fontFamily, copied.FontFamily);
        Assert.Same(shapingProvider, copied.TextShapingProvider);
        Assert.False(copied.UrlPolicy.RestrictUrlSchemes);
    }

    [Fact]
    public void HtmlPdfOptions_DefaultHyperlinkPolicyBlocksFileAndDataUrls() {
        var options = new HtmlPdfSaveOptions();

        Assert.True(options.UrlPolicy.DisallowFileUrls);
        Assert.False(options.UrlPolicy.AllowDataUrls);
        Assert.True(options.UrlPolicy.RestrictUrlSchemes);
        Assert.True(HtmlUrlPolicyEvaluator.IsAllowed("https://example.test/report", options.UrlPolicy));
        Assert.True(HtmlUrlPolicyEvaluator.IsAllowed("mailto:reports@example.test", options.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("cid:report", options.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("ftp://example.test/report", options.UrlPolicy));
        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed("officeimo:report", options.UrlPolicy));
    }

    [Fact]
    public void HtmlPdf_ByteSerializationHonorsCancellation() {
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Cancellation contract</p>").ToPdfDocumentResult();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => {
            _ = HtmlPdfConverterExtensions.SerializeToBytes(result, cancellation.Token);
        });
    }

    [Fact]
    public void MarkdownPdfConverter_RequiresTypedMarkdownDocuments() {
        Assert.DoesNotContain(typeof(MarkdownPdfConverterExtensions).GetMethods(), method =>
            method.GetParameters().Length > 0 && method.GetParameters()[0].ParameterType == typeof(string));
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

    private sealed class NullShapingProvider : OfficeIMO.Drawing.IOfficeTextShapingProvider {
        public OfficeIMO.Drawing.OfficeTextShapingResult? ShapeText(OfficeIMO.Drawing.OfficeTextShapingRequest request) => null;
    }
}
