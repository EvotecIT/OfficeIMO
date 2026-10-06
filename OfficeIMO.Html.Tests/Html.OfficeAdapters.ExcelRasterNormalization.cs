using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlExcelRasterNormalization {
    [Theory]
    [InlineData(HtmlImportMode.Generic, "webp")]
    [InlineData(HtmlImportMode.Semantic, "webp")]
    [InlineData(HtmlImportMode.Generic, "avif")]
    [InlineData(HtmlImportMode.Semantic, "avif")]
    public void ExcelHtml_StaticRasterPreservesNativePicturePixelsAndGeometry(HtmlImportMode mode, string format) {
        string fixture = format == "webp" ? Path.Combine(AppContext.BaseDirectory, "Images", "alpha-gradient")
            : Path.Combine(AppContext.BaseDirectory, "Images", "avif-alpha");
        byte[] sourceBytes = File.ReadAllBytes(fixture + "." + format);
        string source = "data:image/" + format + ";base64," + Convert.ToBase64String(sourceBytes);
        string image = "<img src='" + source + "' width='120' height='80' alt='Independent alpha control'>";
        string html = mode == HtmlImportMode.Generic ? "<p>Before</p>" + image + "<p>After</p>"
            : "<section class='officeimo-sheet' data-officeimo-sheet='Image'><table><tr><td>Before</td></tr></table>"
                + "<section class='officeimo-images'><ul><li data-officeimo-row='3' data-officeimo-column='2'"
                + " data-officeimo-width='120' data-officeimo-height='80'>" + image + "</li></ul></section></section>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = mode });
        using ExcelDocument workbook = result.RequireValue();
        using var output = new MemoryStream();
        workbook.Save(output);
        output.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(output);
        ExcelImage picture = Assert.Single(Assert.Single(reopened.Sheets).Images);
        Assert.Equal("image/png", picture.ContentType);
        Assert.Equal(120, picture.WidthPixels);
        Assert.Equal(80, picture.HeightPixels);
        Assert.True(OfficeRasterImageDecoder.TryDecode(picture.ToBytes(), out OfficeRasterImage? decoded));
        byte[] expected = File.ReadAllBytes(fixture + ".rgba");
        byte[] actual = decoded!.GetPixels();
        Assert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < actual.Length; index++) {
            Assert.InRange(Math.Abs(expected[index] - actual[index]), 0, index % 4 == 3 ? 0 : 3);
        }
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.ResourceTypeUnsupported);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && d.Detail == "sourceMediaType=image/" + format + "; nativeMediaType=image/png");
    }

    [Fact]
    public void ExcelHtml_LayoutRasterUsesTheSameBoundedNativePreparation() {
        string source = "data:image/webp;base64," + Convert.ToBase64String(ReadWebp());
        string html = "<div style='position:absolute;left:20px;top:60px;width:150px;height:100px'>Picture region"
            + "<img src='" + source + "' style='width:120px;height:80px' alt='Positioned control'></div>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.RequireValue();
        ExcelImage picture = Assert.Single(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.True(picture.HasAbsoluteAnchor);
        Assert.Equal("image/png", picture.ContentType);
        Assert.Equal(120, picture.WidthPixels);
        Assert.Equal(80, picture.HeightPixels);
    }

    [Theory]
    [InlineData("pixels", HtmlImportMode.Generic)]
    [InlineData("pixels", HtmlImportMode.Semantic)]
    [InlineData("malformed", HtmlImportMode.Generic)]
    [InlineData("animated", HtmlImportMode.Generic)]
    [InlineData("retainedBytes", HtmlImportMode.Generic)]
    public void ExcelHtml_RejectedNormalizedImageDoesNotConsumeTheNextImageBudget(string reason, HtmlImportMode mode) {
        byte[] source = ReadWebp();
        if (reason == "malformed") source = source.Take(source.Length - 12).ToArray();
        if (reason == "animated") source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "ExcelRaster", "animated.webp"));
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxShapes = 1;
        if (reason == "pixels") limits.MaxDecodedImagePixels = 49 * 33 - 1;
        if (reason == "retainedBytes") {
            Assert.True(OfficeImagePngConverter.TryConvertToPng(source, out byte[] normalized));
            Assert.True(normalized.Length > source.Length);
            limits.MaxImageBytes = normalized.Length - 1;
        }
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black));
        string rejected = "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "'>";
        string retained = "<img src='data:image/png;base64," + Convert.ToBase64String(png) + "' alt='Retained neighbor'>";
        string html = mode == HtmlImportMode.Generic ? rejected + retained
            : "<section class='officeimo-sheet' data-officeimo-sheet='Budget'><table><tr><td>Values</td></tr></table>"
                + "<section class='officeimo-images'><ul><li>" + rejected + "</li><li>" + retained + "</li></ul></section></section>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = mode, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        ExcelImage picture = Assert.Single(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Equal(png, picture.ToBytes());
        Assert.Equal(1, result.Images);
        string code = reason is "pixels" or "retainedBytes" ? HtmlConversionDiagnosticCodes.TargetLimitExceeded
            : HtmlConversionDiagnosticCodes.ResourceDecodeFailed;
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == code
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void ExcelHtml_TotalImageBytesChargesEachNormalizedNativePayload() {
        byte[] source = ReadWebp();
        Assert.True(OfficeImagePngConverter.TryConvertToPng(source, out byte[] normalized));
        Assert.True(normalized.Length > source.Length);
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImageBytes = normalized.Length;
        limits.MaxTotalImageBytes = normalized.Length + source.Length;
        string image = "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "'>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(image + image).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        Assert.Single(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && diagnostic.Detail!.StartsWith("MaxTotalImageBytes:", StringComparison.Ordinal));
    }

    [Fact]
    public void ExcelHtml_TotalImageBytesStillBoundsLargeSourcePayloadsAfterNormalization() {
        byte[] original = ReadWebp();
        var source = new byte[original.Length + 8 + 16384];
        Buffer.BlockCopy(original, 0, source, 0, original.Length);
        Buffer.BlockCopy(System.Text.Encoding.ASCII.GetBytes("JUNK"), 0, source, original.Length, 4);
        Write32(source, original.Length + 4, 16384);
        Write32(source, 4, source.Length - 8);
        Assert.True(OfficeImagePngConverter.TryConvertToPng(source, out byte[] normalized));
        Assert.True(normalized.Length < source.Length);
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImageBytes = source.Length;
        limits.MaxTotalImageBytes = source.Length * 2 - 1;
        string image = "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "'>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(image + image).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        Assert.Single(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Contains(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && d.Detail!.StartsWith("MaxTotalImageBytes:", StringComparison.Ordinal));
    }

    private static void Write32(byte[] bytes, int offset, int value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8));
    }

    private static byte[] ReadWebp() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "alpha-gradient.webp"));
}
