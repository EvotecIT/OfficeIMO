using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelHeaderFooterSvgImageExportTests {
    private const string ColorsSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='16'>"
        + "<rect width='16' height='16' fill='red'/><rect x='16' width='16' height='16' fill='blue'/></svg>";
    private const string UnsupportedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='16'>"
        + "<foreignObject width='32' height='16'/></svg>";

    [Fact]
    public void PersistedSvgHeaderAndFooterMatchPngControl() {
        byte[] svg = Encoding.UTF8.GetBytes(ColorsSvg);
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int loss));
        Assert.Equal(0, loss);
        byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing!);
        OfficeImageExportResult vector = ExportPersisted(svg, "image/svg+xml");
        OfficeImageExportResult raster = ExportPersisted(png, "image/png");

        Assert.DoesNotContain(vector.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        Assert.Equal(raster.Bytes, vector.Bytes);
        Assert.True(OfficePngReader.TryDecode(vector.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(12, 10));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(32, 10));
        Assert.Equal(OfficeColor.Red, image.GetPixel(image.Width / 2 - 8, image.Height - 18));
    }

    [Fact]
    public void SvgHeaderRespectsExistingSectionClip() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='96' height='16'><rect width='96' height='16' fill='red'/></svg>");
        sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left, svg, "image/svg+xml", widthPoints: 72, heightPoints: 12);

        OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png,
            new ExcelWorksheetImageExportOptions {
                Range = "A1:D4", ShowGridlines = false, SplitByManualPageBreaks = true,
                TextShapingProvider = new RecordingShaper(), TextShapingLanguage = "ar-SA"
            })[0];

        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(12, 10));
        Assert.Equal(OfficeColor.White, image.GetPixel(90, 10));
        Assert.Equal(OfficeColor.White, image.GetPixel(12, 30));
    }

    [Fact]
    public void UnsupportedHeaderFooterSvgRetainsFallbackWithSectionSources() {
        OfficeImageExportResult result = ExportPersisted(Encoding.UTF8.GetBytes(UnsupportedSvg), "image/svg+xml");

        string[] sources = result.Diagnostics.Where(d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback)
            .Select(d => d.Source!).ToArray();
        Assert.Equal(new[] { "Report!headerFooter:header-left-image", "Report!headerFooter:footer-center-image" }, sources);
        Assert.All(result.Diagnostics.Where(d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback),
            d => Assert.Equal(OfficeConversionLossKind.Omission, d.LossKind));
    }

    [Fact]
    public void SvgHeaderAndFooterShareIntermediatePixelLimit() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='48' height='48'><rect width='48' height='48' fill='red'/></svg>");
        foreach (ExcelHeaderFooterPosition position in new[] { ExcelHeaderFooterPosition.Left, ExcelHeaderFooterPosition.Center, ExcelHeaderFooterPosition.Right })
            sheet.SetHeaderImage(position, svg, "image/svg+xml", widthPoints: 36, heightPoints: 36);
        sheet.SetFooterImage(ExcelHeaderFooterPosition.Left, svg, "image/svg+xml", widthPoints: 36, heightPoints: 36);
        var options = new ExcelWorksheetImageExportOptions {
            Range = "A1:A1", ShowGridlines = false, SplitByManualPageBreaks = true, MaximumRasterPixels = 10_000,
            TextShapingProvider = new RecordingShaper(), TextShapingLanguage = "ar-SA"
        };
        sheet.ExportImages(OfficeImageExportFormat.Png, options);
        sheet.SetFooterImage(ExcelHeaderFooterPosition.Center, svg, "image/svg+xml", widthPoints: 36, heightPoints: 36);

        Assert.Throws<OfficeImageExportLimitException>(() => sheet.ExportImages(OfficeImageExportFormat.Png, options));
    }

    [Fact]
    public void HeaderSvgCallerCodecCancellationStopsExportBeforeDelivery() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left, Encoding.UTF8.GetBytes(UnsupportedSvg), "image/svg+xml", 24, 12);
        using var cancellation = new CancellationTokenSource();
        var codec = new CancellingCodec(cancellation);
        var options = new ExcelWorksheetImageExportOptions {
            Range = "A1:A1", SplitByManualPageBreaks = true, ImageCodec = codec
        };
        int deliveries = 0;

        Assert.Throws<OperationCanceledException>(() => sheet.ExportImages(OfficeImageExportFormat.Png,
            _ => deliveries++, options, cancellation.Token));
        Assert.Equal(1, codec.Calls);
        Assert.Equal(0, deliveries);
    }

    [Fact]
    public void HeaderSourcePixelLimitReportsOmissionAndStrictPolicyRejectsIt() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left,
            OfficePngWriter.Encode(new OfficeRasterImage(128, 128, OfficeColor.Red)), "image/png", 12, 12);
        var options = new ExcelWorksheetImageExportOptions {
            Range = "A1:A1", SplitByManualPageBreaks = true, MaximumRasterPixels = 5_000
        };

        OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, options)[0];

        OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics,
            d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted);
        Assert.Equal("Report!headerFooter:header-left-image", diagnostic.Source);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        options.Policy = new OfficeImageExportPolicy { RequireNoOmissions = true };
        Assert.Throws<OfficeImageExportPolicyException>(() => sheet.ExportImages(OfficeImageExportFormat.Png, options));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgHeaderAndFooterRetainCallerTextShapingContext(bool footer) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40'>"
            + "<text x='4' y='24' font-family='Arial' font-size='20'>A</text></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int loss));
        Assert.Equal(0, loss);
        var direct = new RecordingShaper();
        OfficeDrawingRasterRenderer.Render(drawing!, new OfficeDrawingRasterRenderOptions {
            TextShapingProvider = direct, TextShapingLanguage = "ar-SA"
        });
        Assert.Contains("A|ar-SA", direct.Requests);
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.AddImage(1, 1, svg, "image/svg+xml", 80, 40);
        var anchored = new RecordingShaper();
        sheet.Range("A1:D4").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
            ShowGridlines = false, TextShapingProvider = anchored, TextShapingLanguage = "ar-SA"
        });
        Assert.Contains("A|ar-SA", anchored.Requests);
        if (footer) sheet.SetFooterImage(ExcelHeaderFooterPosition.Left, svg, "image/svg+xml", 60, 30);
        else sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left, svg, "image/svg+xml", 60, 30);
        var composed = new RecordingShaper();

        sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
            Range = "A1:D4", SplitByManualPageBreaks = true, ShowGridlines = false,
            TextShapingProvider = composed, TextShapingLanguage = "ar-SA"
        });

        // The same text appears in the anchored image and in the selected band.
        Assert.Equal(2, composed.Requests.Count(request => request == "A|ar-SA"));
    }

    private static OfficeImageExportResult ExportPersisted(byte[] bytes, string contentType) {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left, bytes, contentType, widthPoints: 24, heightPoints: 12);
            sheet.SetFooterImage(ExcelHeaderFooterPosition.Center, bytes, contentType, widthPoints: 24, heightPoints: 12);
            document.Save();
        }
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
        return reloaded.Sheets[0].ExportImages(OfficeImageExportFormat.Png,
            new ExcelWorksheetImageExportOptions { Range = "A1:D4", ShowGridlines = false, SplitByManualPageBreaks = true })[0];
    }

    private sealed class CancellingCodec(CancellationTokenSource cancellation) : IOfficeRasterImageCodec {
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            cancellation.Cancel();
            image = new OfficeRasterImage(32, 16, OfficeColor.Red);
            return true;
        }
    }

    private sealed class RecordingShaper : IOfficeTextShapingProvider {
        internal List<string> Requests { get; } = new();
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests.Add(request.Text + "|" + request.Language);
            return null;
        }
    }
}
