using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlApiConsistencyTests {
    [Fact]
    public void DirectHtmlOutputs_UseTargetOrDestinationNaming() {
        Type extensions = typeof(HtmlPdfConverterExtensions);

        Assert.Contains(extensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.ToPdf)
            && method.ReturnType == typeof(byte[])
            && method.GetParameters()[0].ParameterType == typeof(string));

        Assert.DoesNotContain(extensions.GetMethods(), method =>
            method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)
            && method.ReturnType == typeof(byte[]));

        Assert.All(
            extensions.GetMethods().Where(method => method.Name == nameof(HtmlPdfConverterExtensions.SaveAsPdf)),
            method => Assert.Contains(method.GetParameters()[1].ParameterType, new[] { typeof(string), typeof(Stream) }));

        Type[] sources = { typeof(string), typeof(HtmlConversionDocument), typeof(Stream) };
        foreach (Type source in sources) {
            Assert.Contains(extensions.GetMethods(), method =>
                method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
                && method.GetParameters()[0].ParameterType == source
                && method.GetParameters()[1].ParameterType == typeof(string));
            Assert.Contains(extensions.GetMethods(), method =>
                method.Name == nameof(HtmlPdfConverterExtensions.TrySaveAsPdfAsync)
                && method.GetParameters()[0].ParameterType == source
                && method.GetParameters()[1].ParameterType == typeof(Stream));
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

        using var pdfSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var pngSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
        using var svgSource = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));

        Assert.NotEmpty(await pdfSource.ToPdfAsync(options));
        Assert.NotEmpty(await pngSource.ToPngAsync(options));
        Assert.StartsWith("<svg", await svgSource.ToSvgAsync(options), StringComparison.Ordinal);
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
}
