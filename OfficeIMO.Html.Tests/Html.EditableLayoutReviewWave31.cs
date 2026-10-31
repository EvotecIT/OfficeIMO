using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutReviewWave31Tests {
    [Fact]
    public void ProjectedWordPictureUsesRenderedCssDimensions() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<img alt='Sized marker' src='" + image + "' style='width:24px;height:18px'></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence();
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);
        using var reader = new StreamReader(package.MainDocumentPart!.GetStream());
        string documentXml = reader.ReadToEnd();
        Assert.Contains("cx=\"228600\" cy=\"171450\"", documentXml, StringComparison.Ordinal);
    }

    [Fact]
    public void ProjectedWordPictureStillReportsMissingAlternativeText() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<img src='" + image + "' style='width:20px;height:20px'></div>";
        var options = new WordHtmlImportOptions { EnableAccessibilityDiagnostics = true };

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence(options);
        using WordDocument document = result.Value;

        Assert.Single(document.TextBoxes);
        Assert.Single(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "accessibility-image-missing-alt");
    }

    [Fact]
    public void ProjectedExcelRegionRetainsSvgPictureAtRenderedDimensions() {
        string image = CreateSvgDataUri();
        string html = "<div style='position:absolute;width:180px;height:70px'>Region"
            + "<img alt='Vector marker' src='" + image + "' style='width:24px;height:18px'></div>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        ExcelImage nativeImage = Assert.Single(Assert.Single(workbook.Sheets).Images);
        Assert.Equal(24, nativeImage.Width.ToPixels());
        Assert.Equal(18, nativeImage.Height.ToPixels());
        Assert.Equal("Vector marker", nativeImage.Description);
    }

    [Fact]
    public void ProjectedPowerPointRegionRetainsSvgPictureAtRenderedDimensions() {
        string image = CreateSvgDataUri();
        string html = "<div style='position:absolute;width:180px;height:70px'>Region"
            + "<img alt='Vector marker' src='" + image + "' style='width:24px;height:18px'></div>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointDocumentResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointDocument presentation = result.Value;

        PowerPointPicture nativeImage = Assert.Single(Assert.Single(presentation.Slides).Pictures);
        Assert.Equal("image/svg+xml", nativeImage.ContentType);
        Assert.Equal(18D, nativeImage.Width.ToPoints(), precision: 3);
        Assert.Equal(13.5D, nativeImage.Height.ToPoints(), precision: 3);
        Assert.Equal("Vector marker", nativeImage.AltText);
    }

    [Fact]
    public void ProjectedSvgObjectFitCropUsesVisibleBoundsAcrossSheetAndSlide() {
        string image = CreateSvgDataUri();
        string html = "<div style='position:absolute;width:180px;height:70px'>Region"
            + "<img alt='Cropped vector' src='" + image
            + "' style='width:20px;height:20px;object-fit:cover'></div>";

        HtmlToExcelResult excelResult = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = excelResult.Value;
        ExcelImage excelImage = Assert.Single(Assert.Single(workbook.Sheets).Images);
        Assert.Equal(20, excelImage.Width.ToPixels());
        Assert.Equal(20, excelImage.Height.ToPixels());
        Assert.Equal(0.25D, excelImage.CropLeftRatio, precision: 3);
        Assert.Equal(0.25D, excelImage.CropRightRatio, precision: 3);

        HtmlToPowerPointResult powerPointResult = HtmlConversionDocument.Parse(html)
            .ToPowerPointDocumentResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointDocument presentation = powerPointResult.Value;
        PowerPointPicture powerPointImage = Assert.Single(Assert.Single(presentation.Slides).Pictures);
        Assert.Equal(15D, powerPointImage.Width.ToPoints(), precision: 3);
        Assert.Equal(15D, powerPointImage.Height.ToPoints(), precision: 3);
        Assert.Equal(0.25D, powerPointImage.CropLeftRatio, precision: 3);
        Assert.Equal(0.25D, powerPointImage.CropRightRatio, precision: 3);
    }

    private static string CreateSvgDataUri() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<rect width='40' height='20' fill='red'/></svg>";
        return "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
    }
}