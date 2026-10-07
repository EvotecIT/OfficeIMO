using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSvgImageExportTests {
    private const string SupportedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='16'>"
        + "<rect width='16' height='16' fill='red'/><rect x='16' width='16' height='16' fill='blue'/></svg>";
    private const string UnsupportedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='16'>"
        + "<foreignObject width='32' height='16'><div xmlns='http://www.w3.org/1999/xhtml'>Label</div></foreignObject></svg>";

    [Fact]
    public void PersistedSupportedSvgRendersWithoutDecodeFallback() {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            AddImage(document, SupportedSvg);
            document.Save();
        }
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));

        OfficeImageExportResult result = reloaded.Sheets[0].Range("A1:A1").ExportImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ShowGridlines = false });

        Assert.DoesNotContain(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(8, 8));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(24, 8));
    }

    [Fact]
    public void UnsupportedSvgRetainsVisibleFallbackAndImageDiagnostic() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        AddImage(document, UnsupportedSvg);

        OfficeImageExportResult result = document.Sheets[0].Range("A1:A1").ExportImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ShowGridlines = false });

        OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics,
            d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        Assert.Equal("Svg!Probe", diagnostic.Source);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.NotEqual(OfficeColor.White, image!.GetPixel(8, 8));
    }

    [Fact]
    public void UnsupportedSvgCanUseCallerCodec() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        AddImage(document, UnsupportedSvg);
        var codec = new SolidCodec();

        OfficeImageExportResult result = document.Sheets[0].Range("A1:A1").ExportImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ShowGridlines = false, ImageCodec = codec });

        Assert.Equal(1, codec.Calls);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(8, 8));
        Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
    }

    [Fact]
    public void SvgRasterizationSharesIntermediatePixelBudgetAcrossImages() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        AddImage(document, SupportedSvg);
        var options = new ExcelImageExportOptions { ShowGridlines = false, MaximumRasterPixels = 1_500 };
        document.Sheets[0].Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, options);
        document.Sheets[0].AddImage(1, 1, Encoding.UTF8.GetBytes(SupportedSvg), "image/svg+xml", 32, 16, name: "Second probe");
        document.Sheets[0].AddImage(1, 1, Encoding.UTF8.GetBytes(SupportedSvg), "image/svg+xml", 32, 16, name: "Third probe");

        Assert.Throws<OfficeImageExportLimitException>(() => document.Sheets[0].Range("A1:A1")
            .ExportImage(OfficeImageExportFormat.Png, options));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CallerRasterCodecUsesSharedContainerInspection(bool validContainer) {
        // Independently encoded 16x16 JPEG-compressed TIFF, also used
        // by the Core decoder contracts. The JPEG fixture has a size header but no scan.
        byte[] bytes = validContainer
            ? Convert.FromBase64String("SUkqADwAAAD/2P/AABEIABAAEANSEQBHEQBCEQD/2gAMA1IARwBCAAA/APf6+f6+f6KKKKKKKKK//9kACwAAAQMAAQAAABAAAAABAQMAAQAAABAAAAACAQMAAwAAAMYAAAADAQMAAQAAAAcAAAAGAQMAAQAAAAIAAAARAQQAAQAAAAgAAAAVAQMAAQAAAAMAAAAWAQMAAQAAABAAAAAXAQQAAQAAADMAAAAcAQMAAQAAAAEAAABbAQcAIQEAAMwAAAAAAAAACAAIAAgA/9j/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/xAAfAAABBQEBAQEBAQAAAAAAAAAAAQIDBAUGBwgJCgv/xAC1EAACAQMDAgQDBQUEBAAAAX0BAgMABBEFEiExQQYTUWEHInEUMoGRoQgjQrHBFVLR8CQzYnKCCQoWFxgZGiUmJygpKjQ1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4eLj5OXm5+jp6vHy8/T19vf4+fr/2Q==")
            : new byte[] { 0xFF, 0xD8, 0xFF, 0xC0, 0, 17, 8, 0, 1, 0, 1, 3, 1, 17, 0, 2, 17, 0, 3, 17, 0, 0xFF, 0xD9 };
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Inspected");
        sheet.AddImage(1, 1, bytes, validContainer ? "image/tiff" : "image/jpeg", 16, 16, name: "Probe");
        var codec = new SolidCodec(16, 16);

        OfficeImageExportResult result = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ShowGridlines = false, ImageCodec = codec });

        Assert.Equal(validContainer ? 1 : 0, codec.Calls);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(validContainer ? OfficeColor.Red : OfficeColor.White, image!.GetPixel(8, 8));
        Assert.Contains(result.Diagnostics, d => d.Source == "Inspected!Probe" && d.Code == (validContainer
            ? OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec
            : OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizedSourceImageReportsOmissionAndStrictPolicyRejectsIt(bool callerDecodedSvg) {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Limited");
        byte[] bytes = callerDecodedSvg
            ? Encoding.UTF8.GetBytes(UnsupportedSvg)
            : OfficePngWriter.Encode(new OfficeRasterImage(64, 64, OfficeColor.Red));
        sheet.AddImage(1, 1, bytes, callerDecodedSvg ? "image/svg+xml" : "image/png", 16, 16, name: "Oversized");
        var codec = new SolidCodec(64, 64);
        var options = new ExcelImageExportOptions {
            ShowGridlines = false, MaximumRasterPixels = 1_500, ImageCodec = codec
        };

        OfficeImageExportResult result = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(callerDecodedSvg ? 1 : 0, codec.Calls);
        OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics,
            d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted);
        Assert.Equal("Limited!Oversized", diagnostic.Source);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.White, image!.GetPixel(8, 8));

        options.Policy = new OfficeImageExportPolicy { RequireNoOmissions = true };
        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
            sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, options));
        Assert.Contains(exception.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted &&
            d.Source == "Limited!Oversized" && d.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(callerDecodedSvg ? 2 : 0, codec.Calls);
    }

    private static void AddImage(ExcelDocument document, string svg) {
        ExcelSheet sheet = document.AddWorksheet("Svg");
        sheet.AddImage(1, 1, Encoding.UTF8.GetBytes(svg), "image/svg+xml", 32, 16, name: "Probe");
    }

    private sealed class SolidCodec : IOfficeRasterImageCodec {
        private readonly int _width;
        private readonly int _height;
        internal SolidCodec(int width = 32, int height = 16) { _width = width; _height = height; }
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            image = new OfficeRasterImage(_width, _height, OfficeColor.Red);
            return true;
        }
    }
}
