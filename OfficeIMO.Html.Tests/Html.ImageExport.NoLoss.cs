using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task HtmlImageExport_StrictPoliciesRejectTheSameTransformLossBeforeWriting(OfficeImageExportFormat format) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div style='width:40px;height:20px;background:red;transform:perspective(200px) rotateY(30deg)'></div>");
        var options = new HtmlRenderOptions { ViewportWidth = 64D, Margins = HtmlRenderMargins.All(0D) };
        OfficeImageExportResult permissive = document.ExportImage(format, options);
        Assert.Contains(permissive.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        var pdf = document.ToPdfDocumentResult(new HtmlToPdfOptions(options));
        Assert.Contains(pdf.Report.Warnings, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.True(pdf.HasLoss);
        options.Policy.RequireNoLoss = true;
        using var output = new MemoryStream();
        Assert.Throws<OfficeImageExportPolicyException>(() => document.ToImage(options).As(format).Save(output));
        Assert.Equal(0, output.Length);
        await Assert.ThrowsAsync<OfficeImageExportPolicyException>(() => document.ExportImageAsync(format, options));
        options.Policy.RequireNoLoss = false;
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        Assert.Throws<HtmlConversionException>(() => document.ExportImage(format, options));
    }

    [Theory]
    [InlineData("<div style='opacity:invalid;width:20px;height:20px;background:red'></div>")]
    [InlineData("<style>@page{size:160px 120px;margin:0}@page:left{size:100px 120px}</style><div style='height:260px;background:red'></div>")]
    [InlineData("<img src='https://example.invalid/unresolved.png'>")]
    public void HtmlImageExport_RenderWarningsCarryLossThroughEveryReport(string html) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        Assert.True(rendered.HasLoss);
        Assert.All(rendered.Diagnostics.Where(d => d.Severity != HtmlDiagnosticSeverity.Info), diagnostic =>
            Assert.NotEqual(OfficeConversionLossKind.None, diagnostic.LossKind));
        options.Policy.RequireNoLoss = true;
        Assert.Throws<OfficeImageExportPolicyException>(() => document.ExportImage(OfficeImageExportFormat.Svg, options));
    }

    [Fact]
    public void HtmlImageExport_InformationalDiagnosticsDoNotInventLoss() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<table></table><div style='width:20px;height:20px;background:red'></div>");
        var options = new HtmlRenderOptions { ViewportWidth = 64D, Margins = HtmlRenderMargins.All(0D) };
        options.Policy.RequireNoLoss = true;
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.EmptyTable);
        Assert.All(result.Diagnostics, diagnostic => Assert.Equal(OfficeConversionLossKind.None, diagnostic.LossKind));
    }

    [Theory]
    [InlineData("<img src='https://example.invalid/missing.png'>", HtmlRenderDiagnosticCodes.ExternalImagePending)]
    [InlineData("<link rel='stylesheet' href='https://example.invalid/missing.css'><div></div>", HtmlRenderDiagnosticCodes.ExternalStylesheetPending)]
    public void HtmlImageExport_NoOmissionsRejectsUnloadedSourceResources(string html, string code) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions();
        options.Policy.RequireNoOmissions = true;
        OfficeImageExportPolicyException error = Assert.Throws<OfficeImageExportPolicyException>(() =>
            document.ExportImage(OfficeImageExportFormat.Svg, options));
        Assert.Contains(error.Diagnostics, diagnostic => diagnostic.Code == code
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}
