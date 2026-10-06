using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlExcelRasterNormalization {
    [Theory]
    [InlineData("pixels")]
    [InlineData("attempts")]
    public void ExcelHtml_OrientedLayoutNormalizationUsesTheImportWorkBudget(string limit) {
        byte[] exif = Convert.FromBase64String("TU0AKgAAAAgAAQESAAMAAAABAAYAAAAAAAA=");
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeJpegEncodeOptions { Metadata = new OfficeJpegMetadata(exif: exif) });
        var limits = HtmlImportLimits.CreateDefault();
        if (limit == "pixels") limits.MaxDecodedImagePixels = 1;
        else limits.MaxImages = 1;
        string image = "<img src='data:image/jpeg;base64," + Convert.ToBase64String(jpeg) + "' width='120' height='80'>";
        string html = "<div style='position:absolute;left:20px;top:60px;width:300px;height:100px'>Control"
            + image + image + "</div>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        if (limit == "pixels") Assert.Empty(workbook.Sheets.SelectMany(sheet => sheet.Images));
        else Assert.Single(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Contains(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && d.Detail!.Contains(limit == "pixels" ? "MaxDecodedImagePixels" : "raster decode attempts"));
    }

    [Fact]
    public void ExcelHtml_OrientedMultipageTiffLayoutIsOmittedWithoutLosingTheNativeNeighbor() {
        byte[] tiff = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeRasterImage(3, 3, OfficeColor.Blue)
        });
        int ifd = BitConverter.ToInt32(tiff, 4);
        int entries = BitConverter.ToUInt16(tiff, ifd);
        bool changed = false;
        for (int index = 0; index < entries; index++) {
            int entry = ifd + 2 + index * 12;
            if (BitConverter.ToUInt16(tiff, entry) != 274) continue;
            tiff[entry + 8] = 6;
            changed = true;
        }
        Assert.True(changed);
        Assert.True(OfficeRasterContainerInspector.TryInspect(tiff, out OfficeRasterContainerInfo? original));
        Assert.Equal(2, original!.Frames.Count);
        Assert.False(OfficeRasterContainerInspector.TryInspect(tiff,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 2 }, out _));
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black));
        string rejected = "<img src='data:image/tiff;base64," + Convert.ToBase64String(tiff) + "'>";
        string retained = "<img src='data:image/png;base64," + Convert.ToBase64String(png) + "'>";
        string html = "<div style='position:absolute;left:20px;top:60px;width:300px;height:100px'>Control"
            + rejected + retained + "</div>";
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxDecodedImagePixels = 2;
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        using var output = new MemoryStream();
        workbook.Save(output);
        output.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(output);
        Assert.Equal(png, Assert.Single(reopened.Sheets.SelectMany(sheet => sheet.Images)).ToBytes());
        Assert.Contains(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.ResourceDecodeFailed
            && d.LossKind == OfficeConversionLossKind.Omission);
    }

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
    public void ExcelHtml_FailedNormalizationStillConsumesOperationDecodeAdmission() {
        byte[] source = ReadWebp();
        Assert.True(OfficeImagePngConverter.TryConvertToPng(source, out byte[] png));
        Assert.True(png.Length > source.Length);
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxImageBytes = png.Length - 1;
        string image = "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "'>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(image + image + image).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        Assert.Empty(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Equal(2, result.Report.Diagnostics.Count(d => d.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && d.Detail!.Contains("raster decode attempts")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcelHtml_OrientedPositionedWebpRetainsOriginalImportPolicy(bool animated) {
        // Independent libwebp/Pillow EXIF orientation control (3 x 2, rotate clockwise).
        byte[] source = animated ? File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "ExcelRaster", "animated.webp")) : Convert.FromBase64String(
            "UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAGAAAAAAAA");
        // Add the same EXIF orientation to the independent animated control.
        if (animated) {
            byte[] exif = Convert.FromBase64String("TU0AKgAAAAgAAQESAAMAAAABAAYAAAAAAAA=");
            int padded = exif.Length + (exif.Length & 1);
            int offset = source.Length;
            Array.Resize(ref source, offset + 8 + padded);
            Buffer.BlockCopy(System.Text.Encoding.ASCII.GetBytes("EXIF"), 0, source, offset, 4);
            Write32(source, offset + 4, exif.Length);
            Buffer.BlockCopy(exif, 0, source, offset + 8, exif.Length);
            Write32(source, 4, source.Length - 8);
            source[20] |= 8;
        }
        Assert.False(OfficeImageOrientationNormalizer.TryNormalizeToPng(source, out _, out _));
        var limits = HtmlImportLimits.CreateDefault();
        if (!animated) limits.MaxDecodedImagePixels = 5;
        string html = "<div style='position:absolute;left:20px;top:60px;width:150px;height:100px'>Control"
            + "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "' width='120' height='80'></div>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        Assert.Empty(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Contains(result.Report.Diagnostics, d => d.Code == (animated
            ? HtmlConversionDiagnosticCodes.ResourceDecodeFailed : HtmlConversionDiagnosticCodes.TargetLimitExceeded));
    }

    [Fact]
    public void ExcelHtml_FailedNormalizationStillConsumesTotalDecodeInputBytes() {
        byte[] source = ReadWebp();
        Assert.True(OfficeImagePngConverter.TryConvertToPng(source, out byte[] png));
        Assert.True(png.Length > source.Length);
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImageBytes = png.Length - 1;
        limits.MaxTotalImageBytes = limits.MaxImageBytes;
        int admitted = (int)(limits.MaxTotalImageBytes / source.Length);
        string image = "<img src='data:image/webp;base64," + Convert.ToBase64String(source) + "'>";
        string html = string.Concat(Enumerable.Repeat(image, admitted + 1));
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.RequireValue();
        Assert.Empty(workbook.Sheets.SelectMany(sheet => sheet.Images));
        Assert.Single(result.Report.Diagnostics.Where(d => d.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && d.Detail!.Contains("raster decode input bytes")));
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
