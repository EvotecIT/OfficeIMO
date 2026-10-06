using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlExcelPictureHyperlinkTests {
    [Theory]
    [InlineData("https://example.org/photo?item=1&view=full")]
    [InlineData("../photo.html")]
    public void GenericLinkedPhotoRetainsNativeTargetPayloadAndGeometry(string target) {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.SteelBlue));
        string html = "<a href='" + target.Replace("&", "&amp;") + "'><img alt='Photo' width='32' height='16' src='data:image/png;base64," + Convert.ToBase64String(png) + "'></a>";
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument document = result.RequireValue();
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == HtmlConversionDiagnosticCodes.ContentOmitted && item.Message.Contains("hyperlink"));
        using var artifact = document.ToStream();
        using var reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));
        ExcelImage image = reopened.Sheets.Single().Images.Single();
        Assert.Equal(target, image.HyperlinkUri!.OriginalString);
        Assert.Equal(png, image.ToBytes());
        Assert.Equal(32, image.WidthPixels);
        Assert.Equal(16, image.HeightPixels);
        using var native = SpreadsheetDocument.Open(new MemoryStream(artifact.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(native));
    }

    [Fact]
    public void SemanticPictureRoundTripPreservesInertTargetAndExactImage() {
        using var source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Pictures");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        ExcelImage sourceImage = sheet.AddImage(2, 1, png, "image/png", 32, 16);
        var target = new Uri("https://example.org/photo?item=1&view=full");
        sourceImage.HyperlinkUri = target;
        string html = source.ToHtml();
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument restored = result.RequireValue();
        using var artifact = restored.ToStream();
        using var reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));
        ExcelImage image = reopened.Sheets.Single().Images.Single();
        Assert.Equal(target, image.HyperlinkUri);
        Assert.Equal(png, image.ToBytes());
        Assert.Equal(2, image.RowIndex);
        Assert.Equal(1, image.ColumnIndex);
        Assert.Equal(32, image.WidthPixels);
        Assert.Equal(16, image.HeightPixels);
    }
    [Theory]
    [InlineData("https://example.org/photo?item=1&view=full", true)]
    [InlineData("../photo.html", true)]
    [InlineData("javascript:alert(1)", false)]
    public void VisualReviewRetainsSafePictureTargetsAndReportsUnsafeOmission(string target, bool interactive) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Pictures");
        sheet.Cell(1, 1, "Photo");
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        sheet.AddImage(1, 1, png, "image/png", 32, 16).HyperlinkUri = new Uri(target, UriKind.RelativeOrAbsolute);
        var result = sheet.ToHtmlResult(new ExcelHtmlSaveOptions { ExportProfile = ExcelHtmlExportProfile.VisualReview });
        var parsed = System.Xml.Linq.XElement.Parse(result.Value.Substring(result.Value.IndexOf("<svg"), result.Value.IndexOf("</svg>") + 6 - result.Value.IndexOf("<svg")));
        var links = parsed.Descendants().Where(element => element.Name.LocalName == "a").ToArray();
        if (interactive) {
            Assert.Equal(target, Assert.Single(links).Attribute("href")!.Value);
            Assert.Contains(links[0].Descendants(), element => element.Name.LocalName == "image");
            Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ImageHyperlinkUnsupported);
        } else {
            Assert.Empty(links);
            Assert.Contains(result.Report.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ImageHyperlinkUnsupported && item.LossKind == OfficeConversionLossKind.Omission);
            Assert.DoesNotContain(target, result.Value);
        }
        Assert.Single(parsed.Descendants().Where(element => element.Name.LocalName == "image"));
        Assert.Equal(png, sheet.Images.Single().ToBytes());
    }

}
