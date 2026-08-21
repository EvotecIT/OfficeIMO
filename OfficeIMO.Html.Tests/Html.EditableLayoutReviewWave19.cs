using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave19Tests {
    [Fact]
    public void ProjectionCannotImportImagesRejectedByTheDocumentResourcePolicy() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:70px'>Projected"
            + "<img alt='Rejected marker' src='" + image + "' style='width:24px;height:18px'></div>";
        var documentOptions = new HtmlConversionDocumentOptions {
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };

        HtmlToExcelResult excelResult = HtmlConversionDocument.Parse(html, documentOptions)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        HtmlToPowerPointResult powerPointResult = HtmlConversionDocument.Parse(html, documentOptions)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = powerPointResult.Value;

        Assert.Empty(Assert.Single(workbook.Sheets).Images);
        Assert.Empty(Assert.Single(presentation.Slides).Pictures);
        Assert.Contains(excelResult.Report.Diagnostics, IsRejectedImageDiagnostic);
        Assert.Contains(powerPointResult.Report.Diagnostics, IsRejectedImageDiagnostic);
    }

    [Fact]
    public void PowerPointSolidFillStaysBehindNativeBackgroundPicture() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style=\"position:absolute;width:180px;height:50px;background-color:#123456;"
            + "background-image:url('" + image + "');background-repeat:no-repeat;background-size:24px 18px\">Layered</div>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointAutoShape backingFill = Assert.Single(slide.Shapes.OfType<PowerPointAutoShape>(),
            shape => shape.FillColor == "123456");
        PowerPointPicture picture = Assert.Single(slide.Pictures);
        PowerPointTextBox textBox = Assert.Single(slide.TextBoxes, box => box.Text == "Layered");
        List<PowerPointShape> shapes = slide.Shapes.ToList();

        Assert.Equal(100, textBox.FillTransparency);
        Assert.True(shapes.IndexOf(backingFill) < shapes.IndexOf(picture));
        Assert.True(shapes.IndexOf(picture) < shapes.IndexOf(textBox));
    }

    [Fact]
    public void PowerPointBackgroundPictureIsOmittedWhenItsBackingFillExceedsTheShapeBudget() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style=\"position:absolute;width:180px;height:50px;background-color:#123456;"
            + "background-image:url('" + image + "');background-repeat:no-repeat;background-size:24px 18px\">Layered</div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 1;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                Limits = limits
            });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.Empty(slide.Pictures);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void EmptyLaterLayoutFragmentsStillKeepPagedRegionsInSemanticFlow() {
        const string html = "<div style='display:flex;width:220px;height:420px'>"
            + "<span style='height:32px'>Only visible fragment</span></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(12D)
        };

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            options,
            HtmlCssMediaContext.Print);

        Assert.Empty(projection.Regions);
        Assert.Contains("Only visible fragment", projection.RemainingDocument.Body!.TextContent,
            StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionFragmented
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("occurrences=", StringComparison.Ordinal));
    }

    private static bool IsRejectedImageDiagnostic(HtmlDiagnostic diagnostic) =>
        diagnostic.Code == "ImageResourceRejectedByPolicy";
}
