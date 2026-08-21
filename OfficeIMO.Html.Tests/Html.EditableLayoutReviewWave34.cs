using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutReviewWave34Tests {
    [Theory]
    [InlineData("img", "absolute")]
    [InlineData("span", "fixed")]
    [InlineData("input", "absolute")]
    public void PositionedNonContainerElementStaysInDiagnosedSemanticFlow(
        string elementName,
        string position) {
        string opening = "<" + elementName + " style='position:" + position
            + ";left:20px;top:10px;width:40px;height:20px'";
        string html = elementName == "img"
            ? opening + " alt='Positioned picture' src='data:image/png;base64,"
                + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2)) + "'>"
            : elementName == "input"
                ? opening + " value='Positioned input'>"
                : opening + ">Positioned inline</" + elementName + ">";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.NotNull(projection.RemainingDocument.QuerySelector(elementName));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "unsupportedPositionedElement=" + elementName + "; semanticFlow=true");
    }

    [Fact]
    public void PositionedNonContainerInsideNativeRegionUsesOwningRegionWithoutDuplicateBoundary() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:100px;height:60px'>"
            + "<img style='position:absolute;left:10px;top:8px;width:20px;height:20px' src='" + image + "'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Single(projection.Regions);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic =>
            diagnostic.Detail == "unsupportedPositionedElement=img; semanticFlow=true");
    }

    [Fact]
    public void PositionedImageBoundaryDiagnosticFlowsThroughEveryEditableDestination() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<img alt='Positioned picture' src='" + image
            + "' style='position:absolute;left:20px;top:10px;width:40px;height:20px'>";

        HtmlToWordResult wordResult = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = wordResult.Value;
        HtmlToRtfResult rtfResult = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        HtmlToExcelResult excelResult = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        HtmlToPowerPointResult powerPointResult = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;

        Assert.Contains(wordResult.Report.Diagnostics, IsUnsupportedPositionedImageDiagnostic);
        Assert.Contains(rtfResult.Report.Diagnostics, IsUnsupportedPositionedImageDiagnostic);
        Assert.Contains(excelResult.Report.Diagnostics, IsUnsupportedPositionedImageDiagnostic);
        Assert.Contains(powerPointResult.Report.Diagnostics, IsUnsupportedPositionedImageDiagnostic);
    }

    [Fact]
    public void RendererRejectedRtfRegionPictureStaysOutOfNativeFrame() {
        int rejectedLength = checked((int)new HtmlRenderOptions().MaxResourceBytes + 1);
        byte[] oversizedPng = new byte[rejectedLength];
        byte[] validPrefix = PdfPngTestImages.CreateRgbPng(2, 2);
        Array.Copy(validPrefix, oversizedPng, validPrefix.Length);
        string image = "data:image/png;base64," + Convert.ToBase64String(oversizedPng);
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<img alt='Budget rejected' src='" + image + "'></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        Assert.DoesNotContain(result.Value.Paragraphs, paragraph => paragraph.Frame.HasAnyValue);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "unrenderedRegionImage=true; semanticFlow=true");
    }

    private static bool IsUnsupportedPositionedImageDiagnostic(HtmlDiagnostic diagnostic) =>
        diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
        && diagnostic.Detail == "unsupportedPositionedElement=img; semanticFlow=true";
}