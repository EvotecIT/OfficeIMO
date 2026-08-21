using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave18Tests {
    [Theory]
    [InlineData("padding:12px")]
    [InlineData("padding-top:3px;padding-right:4px;padding-bottom:5px;padding-left:6px")]
    public void PaddedRegionsStayInDiagnosedSemanticFlow(string padding) {
        string html = "<div style='position:absolute;width:180px;height:50px;" + padding + "'>Padded</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Padded", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("padding", StringComparison.Ordinal)
            && diagnostic.Detail.Contains("semanticFlow=true", StringComparison.Ordinal));
    }

    [Fact]
    public void DescendantPaddingKeepsOwningRegionInSemanticFlow() {
        const string html = "<div style='position:absolute;width:180px;height:50px'>"
            + "<span style='padding:8px'>Padded child</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("descendant:padding=8px", StringComparison.Ordinal));
    }

    [Fact]
    public void PositionedImageChildrenStaySemanticForWordAndRtf() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<img alt='Nested marker' src='" + image
            + "' style='position:absolute;left:28px;top:18px;width:24px;height:18px'></div>";

        HtmlToWordResult wordResult = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = wordResult.Value;
        HtmlToRtfResult rtfResult = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        string rtf = rtfResult.Value.ToRtf();

        Assert.Empty(word.TextBoxes);
        Assert.Single(word.Images);
        Assert.DoesNotContain(@"\phpg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pict", rtf, StringComparison.Ordinal);
        Assert.Contains(wordResult.Report.Diagnostics, IsNestedPlacementDiagnostic);
        Assert.Contains(rtfResult.Report.Diagnostics, IsNestedPlacementDiagnostic);
    }

    [Fact]
    public void NonImagePowerPointBackgroundLayersAreDiagnosedAsOmitted() {
        const string html = "<div style='position:absolute;width:180px;height:50px;"
            + "background-image:linear-gradient(red,blue)'>Gradient</div>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Single(Assert.Single(presentation.Slides).TextBoxes, box => box.Text == "Gradient");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened
            && diagnostic.Severity == HtmlDiagnosticSeverity.Warning
            && diagnostic.LossKind == OfficeConversionLossKind.Omission
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("retainedNativePictures=0", StringComparison.Ordinal)
            && diagnostic.Detail.Contains("omittedLayers=1", StringComparison.Ordinal));
    }

    private static bool IsNestedPlacementDiagnostic(HtmlDiagnostic diagnostic) =>
        diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
        && diagnostic.Detail == "nestedLayoutPlacement=true; semanticFlow=true";
}
