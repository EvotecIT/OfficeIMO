using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave23Tests {
    [Fact]
    public void ExcelAndPowerPointKeepMixedInlinePicturesInSemanticFlow() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:50px'>Before"
            + "<img alt='Middle' src='" + image + "'>After</div>";

        HtmlToExcelResult excelResult = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        HtmlToPowerPointResult powerPointResult = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;

        Assert.Contains(excelResult.Report.Diagnostics, IsMixedInlineRetentionDiagnostic);
        Assert.Contains(powerPointResult.Report.Diagnostics, IsMixedInlineRetentionDiagnostic);
    }

    [Theory]
    [InlineData("margin:12px")]
    [InlineData("margin-left:80px")]
    [InlineData("margin-top:3rem")]
    public void DescendantMarginsRemainInSemanticFlow(string margin) {
        string html = "<div style='position:absolute;width:180px;height:50px'>"
            + "<div style='" + margin + "'>Inset</div></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("margin", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("<div lang='fr'>Bonjour</div>")]
    [InlineData("<span xml:lang='de-DE'>Guten Tag</span>")]
    public void LanguageScopedTextRemainsInSemanticFlow(string content) {
        string html = "<div style='position:absolute;width:180px;height:50px'>"
            + content + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, IsSemanticRetentionDiagnostic);
    }

    [Theory]
    [InlineData("<ruby>Base<rt>Annotation</rt></ruby>")]
    [InlineData("<math><mfrac><mi>x</mi><mn>2</mn></mfrac></math>")]
    public void RubyAndMathMlRemainInSemanticFlow(string content) {
        string html = "<div style='position:absolute;width:180px;height:50px'>"
            + content + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, IsSemanticRetentionDiagnostic);
    }

    [Fact]
    public void GeneratedPseudoElementContentRemainsInSemanticFlow() {
        const string html = "<style>.region::before{content:'Badge';display:block;"
            + "background:red;border:2px solid black}</style>"
            + "<div class='region' style='position:absolute;width:180px;height:50px'>Body</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "generatedContent=true; semanticFlow=true");
    }

    private static bool IsMixedInlineRetentionDiagnostic(HtmlDiagnostic diagnostic) =>
        diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
        && diagnostic.Detail == "mixedInlinePictures=true";

    private static bool IsSemanticRetentionDiagnostic(HtmlDiagnostic diagnostic) =>
        diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
        && diagnostic.Detail == "semanticContent=true";
}
