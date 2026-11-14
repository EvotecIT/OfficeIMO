using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

/// <summary>Format and attribute choices the printer service makes for an IPP printer, from its Get-Printer-Attributes answer.</summary>
public sealed class IppPrinterDeliveryTests {
    private static readonly IppPrinterCapabilities PdfAndRaster = new() {
        DocumentFormats = ["application/octet-stream", "application/pdf", "image/pwg-raster"],
        Media = ["na_letter_8.5x11in", "iso_a4_210x297mm", "custom_min_3x5in", "custom_max_8.5x14in"],
        MediaDefault = "na_letter_8.5x11in",
        MediaSizes = [new IppMedia(21590, 27940) { SizeName = "na_letter_8.5x11in" }, new IppMedia(21000, 29700) { SizeName = "iso_a4_210x297mm" }],
        MediaSources = ["auto", "main", "by-pass-tray"],
        Sides = ["one-sided", "two-sided-long-edge", "two-sided-short-edge"],
        ColorModes = ["auto", "color", "monochrome"],
        MaxCopies = 99,
        Orientations = [IppOrientation.Portrait, IppOrientation.Landscape],
        PrintScaling = ["auto", "fit", "none"],
        PwgRasterResolutions = [new(150, 150, IppResolutionUnits.DotsPerInch), new(300, 300, IppResolutionUnits.DotsPerInch), new(600, 600, IppResolutionUnits.DotsPerInch)],
        PwgRasterTypes = ["black_1", "sgray_8", "srgb_8"],
        Resolutions = [new(118, 118, IppResolutionUnits.DotsPerCentimeter)]
    };

    [Fact]
    public void CapabilitiesMapMediaSourcesAndFeaturesAndSkipCustomRanges() {
        PdfPrinterCapabilities capabilities = IppPdfPrinter.MapCapabilities("ipp://printer/ipp/print", PdfAndRaster);

        Assert.Equal(["na_letter_8.5x11in", "iso_a4_210x297mm"], capabilities.Media.Select(media => media.Id));
        Assert.Equal(612, capabilities.DefaultMedia!.Size!.Value.Width.ToPoints(), 1);
        Assert.Equal(["main", "by-pass-tray"], capabilities.PaperSources.Select(source => source.Id));
        Assert.True(capabilities.SupportsDuplex);
        Assert.True(capabilities.SupportsColor);
        Assert.Equal(99, capabilities.MaximumCopies);
        Assert.Equal(new PdfPrinterResolution(300, 300), Assert.Single(capabilities.Resolutions));
    }

    [Fact]
    public void ColorVectorJobsSendPdfWithMediaAndNoScaling() {
        IppDeliveryPlan plan = Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 2), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", Copies = 3, Duplex = PdfPrintDuplex.LongEdge }, PdfAndRaster);

        Assert.Equal(IppDocumentKind.VectorPdf, plan.Kind);
        Assert.Equal("application/pdf", plan.DocumentFormat);
        Assert.Equal(3, plan.JobOptions.Copies);
        Assert.Equal("two-sided-long-edge", plan.JobOptions.Sides);
        Assert.Equal("iso_a4_210x297mm", plan.JobOptions.Media);
        Assert.Null(plan.JobOptions.Orientation); // the composed PDF pages carry their own orientation
        Assert.Equal("none", plan.JobOptions.PrintScaling);
        Assert.Equal("color", plan.JobOptions.ColorMode);
        Assert.Empty(plan.Ignored);
    }

    [Theory]
    [InlineData(PdfPrintContentKind.Vector, PdfPrintColorMode.Grayscale)]
    [InlineData(PdfPrintContentKind.Raster, PdfPrintColorMode.Color)]
    public void RasterizedGrayscaleAndRasterJobsUsePwgRaster(PdfPrintContentKind kind, PdfPrintColorMode color) {
        IppDeliveryPlan plan = Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 3, color), kind,
            new PdfPrintDeliveryOptions { PrinterName = "p", Copies = 2, MaximumRasterDpi = 400, RasterizeGrayscale = true }, PdfAndRaster);

        Assert.Equal(IppDocumentKind.PwgRaster, plan.Kind);
        Assert.Equal("image/pwg-raster", plan.DocumentFormat);
        Assert.Equal(300, plan.RasterDpi);
        Assert.Equal(color == PdfPrintColorMode.Grayscale ? PwgRasterColorSpace.Gray8 : PwgRasterColorSpace.Rgb8, plan.ColorSpace);
        Assert.Equal(color == PdfPrintColorMode.Grayscale ? "monochrome" : "color", plan.JobOptions.ColorMode);
        // Multi-sheet raster copies are repeated in the document so they stay collated.
        Assert.Null(plan.JobOptions.Copies);
        Assert.Equal(2, plan.SoftwareCopies);
        Assert.Null(plan.JobOptions.Orientation);
        Assert.Equal(["iso_a4_210x297mm", "iso_a4_210x297mm", "iso_a4_210x297mm"], plan.MediaNames);
    }

    [Fact]
    public void PrintersWithoutPwgRasterReceiveRasterPdfAndThoseWithNeitherFormatAreRefused() {
        var pdfOnly = new IppPrinterCapabilities { DocumentFormats = ["application/pdf"], MaxCopies = 1 };
        Assert.Equal(IppDocumentKind.RasterPdf, Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1, PdfPrintColorMode.Grayscale),
            PdfPrintContentKind.Vector, new PdfPrintDeliveryOptions { PrinterName = "p" }, pdfOnly).Kind);
        Assert.Throws<NotSupportedException>(() => Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", Copies = 2 }, pdfOnly));
        // Without media keywords or media-col support, the paper size cannot be requested and is reported.
        Assert.Equal(["media"], Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p" }, pdfOnly).Ignored);
        var urfOnly = new IppPrinterCapabilities { DocumentFormats = ["image/urf"] };
        Assert.Throws<NotSupportedException>(() => Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p" }, urfOnly));
    }

    [Fact]
    public void UnsupportedSidesAreReportedAndPaperSourcesUseMediaCol() {
        var simplex = new IppPrinterCapabilities {
            DocumentFormats = ["application/pdf"], Sides = ["one-sided"], MediaSources = ["main"], Media = ["iso_a4_210x297mm"]
        };
        IppDeliveryPlan plan = Plan(Sheets(PdfPageSizes.A4.Height, PdfPageSizes.A4.Width, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", Duplex = PdfPrintDuplex.ShortEdge, PaperSourceId = "main" }, simplex);

        Assert.Equal(["sides"], plan.Ignored);
        Assert.Null(plan.JobOptions.Sides);
        Assert.Null(plan.JobOptions.Media);
        Assert.Equal(new IppMedia(21000, 29700) { SizeName = "iso_a4_210x297mm", Source = "main" }, plan.JobOptions.MediaCol);
        Assert.Throws<ArgumentException>(() => Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", PaperSourceId = "by-pass-tray" }, simplex));
    }

    [Fact]
    public void UnlistedSizesAreRequestedByDimensionsAndMixedSizesLeaveMediaToEachPage() {
        var custom = new PdfPageSize(OfficeLength.Points(250), OfficeLength.Points(400));
        IppDeliveryPlan plan = Plan(Sheets(custom.Width, custom.Height, 1), PdfPrintContentKind.Vector, new PdfPrintDeliveryOptions { PrinterName = "p" }, PdfAndRaster);
        Assert.Equal(new IppMedia(8819, 14111), plan.JobOptions.MediaCol);

        var mixed = new PdfPrintPlan(2, [1, 2], [new PdfPrintSheet(1, PdfPageSizes.A4, []), new PdfPrintSheet(2, PdfPageSizes.Letter, [])]);
        plan = Plan(mixed, PdfPrintContentKind.Vector, new PdfPrintDeliveryOptions { PrinterName = "p" }, PdfAndRaster);
        Assert.Null(plan.JobOptions.Media);
        Assert.Null(plan.JobOptions.MediaCol);
        Assert.Equal(["iso_a4_210x297mm", "na_letter_8.5x11in"], plan.MediaNames);
    }

    [Fact]
    public void SoftwareCopiesFollowThePrinterDefaultSidesWhenTheJobOmitsThem() {
        var duplexDefault = new IppPrinterCapabilities {
            DocumentFormats = PdfAndRaster.DocumentFormats, Media = PdfAndRaster.Media, MediaSizes = PdfAndRaster.MediaSizes, MaxCopies = 99,
            Sides = PdfAndRaster.Sides, SidesDefault = "two-sided-short-edge",
            PwgRasterResolutions = PdfAndRaster.PwgRasterResolutions, PwgRasterTypes = PdfAndRaster.PwgRasterTypes
        };
        IppDeliveryPlan plan = Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 3, PdfPrintColorMode.Grayscale), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", Copies = 2 }, duplexDefault);

        Assert.Null(plan.JobOptions.Sides);
        Assert.Equal(2, plan.SoftwareCopies);
        Assert.True(plan.Duplex);
        Assert.True(plan.Tumble);
        // The second copy starts on a fresh sheet rather than the back of the first copy's last sheet.
        Assert.Equal(new int?[] { 0, 1, 2, null, 0, 1, 2 }, PwgRasterJobStream.PageSequence(3, plan.SoftwareCopies, plan.Duplex));

        plan = Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 3, PdfPrintColorMode.Grayscale), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", Copies = 2, Duplex = PdfPrintDuplex.SingleSided }, duplexDefault);
        Assert.Equal("one-sided", plan.JobOptions.Sides);
        Assert.False(plan.Duplex);
    }

    [Fact]
    public void DuplexCopiesStartOnANewSheetAndBackSidesFollowTheSheetBackRule() {
        Assert.Equal(new int?[] { 0, 1, 2, null, 0, 1, 2 }, PwgRasterJobStream.PageSequence(3, 2, duplex: true));
        Assert.Equal(new int?[] { 0, 1, 0, 1 }, PwgRasterJobStream.PageSequence(2, 2, duplex: true));
        Assert.Equal((false, true), PwgRasterJobStream.BackSideTransform("flipped", tumble: false));
        Assert.Equal((true, false), PwgRasterJobStream.BackSideTransform("flipped", tumble: true));
        Assert.Equal((true, true), PwgRasterJobStream.BackSideTransform("rotated", tumble: false));
        Assert.Equal((false, false), PwgRasterJobStream.BackSideTransform("normal", tumble: true));
    }

    [Fact]
    public void SoftwareCopiesReuseEachEncodedSheetAndMatchRenderingItAgain() {
        PdfDocument document = PdfDocument.Create(c => {
            foreach (string label in new[] { "One", "Two", "Three" })
                c.Page(p => p.Size(OfficeLength.Points(144), OfficeLength.Points(216)).Content(content => content.AddParagraph(label)));
        });
        PdfPrintJob job = PdfPrintJob.Create(document, new() { InputPath = "copies.pdf", ColorMode = PdfPrintColorMode.Grayscale });
        var plan = new IppDeliveryPlan {
            Kind = IppDocumentKind.PwgRaster, SoftwareCopies = 2, Duplex = true, SheetBack = "rotated", RasterDpi = 72, ColorSpace = PwgRasterColorSpace.Gray8
        };
        using var reused = new PwgRasterJobStream(job, plan, default);
        using var rendered = new PwgRasterJobStream(job, plan, keptPageBytes: 0, default);
        Assert.Equal(ReadAll(rendered), ReadAll(reused));
        Assert.Equal(3, reused.ReusedPages); // 3 sheets x 2 copies render 3 times, not 6.
        Assert.Equal(0, rendered.ReusedPages);

        byte[] pixels = Enumerable.Range(0, 4096).Select(index => (byte)(index % 7 == 0 ? 0 : 255)).ToArray();
        Assert.Equal(pixels, PackedSheetPixels.Pack(pixels, 32, 32, 300).Unpack());
    }

    private static byte[] ReadAll(Stream stream) {
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    [Fact]
    public void LandscapeSheetsTurnToThePortraitFeedWithTheirTopOnTheLeftEdge() {
        // A 3x2 landscape raster: red top-left, blue top-right, the rest white.
        var raster = new OfficeRasterImage(3, 2);
        for (int y = 0; y < 2; y++) for (int x = 0; x < 3; x++) raster.SetPixel(x, y, OfficeColor.White);
        raster.SetPixel(0, 0, OfficeColor.Red);
        raster.SetPixel(2, 0, OfficeColor.Blue);

        byte[] rgb = PwgRasterJobStream.DevicePixels(raster, 3, false, false, 1, out int width, out int height, default);

        Assert.Equal((2, 3), (width, height));
        Assert.Equal(new byte[] { 0, 0, 255 }, rgb.Take(3)); // top-left now holds the top-right corner
        Assert.Equal(new byte[] { 255, 0, 0 }, rgb.Skip((2 * width + 0) * 3).Take(3)); // bottom-left holds the top-left corner
        byte[] mirrored = PwgRasterJobStream.DevicePixels(raster, 1, true, true, 1, out _, out _, default);
        Assert.Equal((byte)((0 * 77 + 0 * 150 + 255 * 29 + 128) >> 8), mirrored[(2 * width + 1)]); // rotated 180 degrees

        // Rendered at half the printer resolution, every device pixel is repeated twice along both edges.
        byte[] doubled = PwgRasterJobStream.DevicePixels(raster, 3, false, false, 2, out int doubledWidth, out int doubledHeight, default);
        Assert.Equal((4, 6), (doubledWidth, doubledHeight));
        foreach ((int x, int y) in new[] { (0, 0), (1, 0), (0, 1), (1, 1) }) Assert.Equal(new byte[] { 0, 0, 255 }, doubled.Skip((y * doubledWidth + x) * 3).Take(3));
        Assert.Equal(new byte[] { 255, 0, 0 }, doubled.Skip((5 * doubledWidth + 1) * 3).Take(3));
    }

    [Theory]
    [InlineData(300, 600, 2)]  // The default limit on a 600 dpi-only laser renders at 300 dpi.
    [InlineData(250, 600, 3)]  // The largest whole fraction within the limit.
    [InlineData(600, 600, 1)]  // A caller asking for 600 dpi gets true 600 dpi.
    [InlineData(72, 600, 1)]   // No fraction up to an eighth fits, so the printer resolution is rendered directly.
    public void PrintersListingOnlyHigherResolutionsRenderAtAFractionWithinTheLimit(double maximumDpi, int printerDpi, int expectedScale) {
        IppPrinterCapabilities printer = new() {
            DocumentFormats = ["image/pwg-raster"], Media = PdfAndRaster.Media, MediaSizes = PdfAndRaster.MediaSizes, MaxCopies = 99,
            Sides = PdfAndRaster.Sides, PwgRasterTypes = PdfAndRaster.PwgRasterTypes,
            PwgRasterResolutions = [new(printerDpi, printerDpi, IppResolutionUnits.DotsPerInch)]
        };
        IppDeliveryPlan plan = Plan(Sheets(PdfPageSizes.A4.Width, PdfPageSizes.A4.Height, 1), PdfPrintContentKind.Vector,
            new PdfPrintDeliveryOptions { PrinterName = "p", MaximumRasterDpi = maximumDpi }, printer);

        Assert.Equal(printerDpi, plan.RasterDpi);
        Assert.Equal(expectedScale, plan.RasterScale);
    }

    private static IppDeliveryPlan Plan(PdfPrintPlan sheets, PdfPrintContentKind kind, PdfPrintDeliveryOptions options, IppPrinterCapabilities printer) =>
        IppPdfPrinter.Plan(sheets, kind, 36_000_000, options, printer, "p");

    private static PdfPrintPlan Sheets(OfficeLength width, OfficeLength height, int count, PdfPrintColorMode color = PdfPrintColorMode.Color) =>
        new(count, Enumerable.Range(1, count).ToArray(),
            Enumerable.Range(1, count).Select(number => new PdfPrintSheet(number, new PdfPageSize(width, height), [])).ToArray(), color);
}
