using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintJobTests {
    [Theory]
    [InlineData(PdfPrintAlignment.TopLeft, false)]
    [InlineData(PdfPrintAlignment.TopRight, true)]
    public void CustomScaleCropsTheChosenSourceEdgeAndKeepsPaperMarginsWhite(PdfPrintAlignment alignment, bool showsRightStripe) {
        OfficeShape background = OfficeShape.Rectangle(100, 100), stripe = OfficeShape.Rectangle(20, 100);
        background.FillColor = OfficeColor.Red; background.StrokeWidth = 0;
        stripe.FillColor = OfficeColor.Blue; stripe.StrokeWidth = 0;
        PdfDocument document = Square(100).Canvas(canvas => canvas.Shape(background, OfficeLength.Points(0), OfficeLength.Points(0))
            .Shape(stripe, OfficeLength.Points(80), OfficeLength.Points(0)));
        PdfPrintJob job = PdfPrintJob.Create(document, new() {
            InputPath = "snapshot.pdf", PaperSize = Paper(150, 180), Margin = 10,
            ScaleMode = PdfPrintScaleMode.Custom, CustomScalePercent = 250, Alignment = alignment
        });
        Assert.Equal(PdfPrintContentKind.Vector, job.ContentKind);
        Assert.True(Assert.Single(Assert.Single(job.Plan.Sheets).Placements).IsClipped);
        OfficeRasterImage raster = Decode(job.RenderSheet(1, 72));
        Assert.Equal(showsRightStripe ? OfficeColor.Blue : OfficeColor.Red, raster.GetPixel(120, 60));
        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 60));
        Assert.Equal(OfficeColor.White, raster.GetPixel(145, 60));
        Assert.Equal(OfficeColor.White, raster.GetPixel(60, 5));
        Assert.Equal(OfficeColor.White, raster.GetPixel(60, 175));
    }

    [Fact]
    public void GrayscaleSheetsPreserveLayoutAndSourceAndContainOnlyNeutralPixels() {
        OfficeShape marker = OfficeShape.Rectangle(20, 20);
        marker.FillColor = OfficeColor.Red;
        marker.StrokeWidth = 0;
        PdfDocument document = Square(100).Canvas(canvas => canvas.Shape(marker, OfficeLength.Points(0), OfficeLength.Points(0)));
        byte[] source = document.ToBytes();
        var request = new PdfPrintPlanRequest {
            InputPath = "snapshot.pdf", PaperSize = Paper(200, 200), Margin = 10,
            ScaleMode = PdfPrintScaleMode.Custom, CustomScalePercent = 50, Alignment = PdfPrintAlignment.BottomRight
        };
        OfficeRasterImage color = Decode(PdfPrintJob.Create(document, request).RenderSheet(1, 72));
        Assert.Equal(OfficeColor.Red, color.GetPixel(145, 145));
        Assert.Equal(OfficeColor.White, color.GetPixel(15, 15));
        request.ColorMode = PdfPrintColorMode.Grayscale;
        PdfPrintJob job = PdfPrintJob.Create(document, request);
        OfficeRasterImage gray = Decode(job.RenderSheet(1, 72));
        Assert.Equal(OfficeColorTransforms.Grayscale(OfficeColor.Red), gray.GetPixel(145, 145));
        for (int y = 0; y < gray.Height; y++) for (int x = 0; x < gray.Width; x++) {
            OfficeColor pixel = gray.GetPixel(x, y);
            Assert.Equal(pixel.R, pixel.G); Assert.Equal(pixel.G, pixel.B);
        }
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void VectorContentKeepsSourceTextAndPlacesEveryPageOfAnNUpSheet() {
        PdfDocument document = PdfDocument.Create(builder => {
            for (int index = 1; index <= 5; index++) {
                int number = index;
                builder.Page(page => page.Size(PdfPageSizes.A4).Content(content =>
                    content.Item(item => item.AddParagraph(text => text.AddText("Source page " + number)))));
            }
        });
        PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest {
            InputPath = "snapshot.pdf", PaperSize = PdfPageSizes.A4, PagesPerSheet = 4, PageSubset = PdfPrintPageSubset.All
        });
        Assert.Equal(PdfPrintContentKind.Vector, job.ContentKind);
        byte[] content = ReadContent(job);
        PdfReadDocument read = PdfReadDocument.Open(content);
        Assert.Equal(2, read.Pages.Count);
        string text = read.ExtractText();
        for (int number = 1; number <= 5; number++) Assert.Contains("Source page " + number, text);
    }

    [Fact]
    public void FilledFormFieldsArePrintedFromVectorContent() {
        byte[] form = PdfDocument.Create(new PdfCreateOptions {
            PageSetup = { Size = Paper(320, 220), Margins = { Left = OfficeLength.Points(36), Right = OfficeLength.Points(36), Top = OfficeLength.Points(36), Bottom = OfficeLength.Points(36) } }
        }).TextField("Customer", width: OfficeLength.Points(160), height: OfficeLength.Points(20), value: "Filled value").ToBytes();
        PdfDocument document = PdfDocument.Load(form);
        PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest { InputPath = "form.pdf", PaperSize = PdfPageSizes.A4 });
        Assert.Equal(PdfPrintContentKind.Vector, job.ContentKind);
        // Imported pages carry content streams only, so the widget appearance must already be part of the page.
        PdfReadDocument read = PdfReadDocument.Open(ReadContent(job));
        Assert.Contains("Filled value", read.ExtractText());
        Assert.Equal(form, document.ToBytes());
    }

    [Fact]
    public void AutomaticOrientationGivesEachContentPageItsOwnPaperOrientation() {
        PdfDocument document = PdfDocument.Create(builder => {
            builder.Page(page => page.Size(OfficeLength.Points(200), OfficeLength.Points(300)));
            builder.Page(page => page.Size(OfficeLength.Points(300), OfficeLength.Points(200)));
        });
        PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = PdfPageSizes.A4 });
        using Stream stream = job.OpenContent();
        IReadOnlyList<PdfPageInfo> pages = PdfInspector.Inspect(stream).Pages;
        Assert.Equal(2, pages.Count);
        Assert.True(pages[0].Height > pages[0].Width);
        Assert.True(pages[1].Width > pages[1].Height);
        Assert.Equal(job.Plan.Sheets[1].PaperSize.Width.ToPoints(), pages[1].Width, 1);
    }

    [Theory]
    [InlineData(PdfPrintScaleMode.Fit)]
    [InlineData(PdfPrintScaleMode.Fill)]
    public void EnlargedVectorDetailIsRasterizedAtTheSheetResolution(PdfPrintScaleMode mode) {
        OfficeShape stripe = OfficeShape.Rectangle(0.2, 72);
        stripe.FillColor = OfficeColor.Black;
        stripe.StrokeWidth = 0;
        PdfDocument document = Square(72).Canvas(canvas => canvas.Shape(stripe, OfficeLength.Points(30), OfficeLength.Points(0)));
        PdfPrintJob job = PdfPrintJob.Create(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = Paper(720, 720), Margin = 0, ScaleMode = mode });
        OfficeRasterImage raster = Decode(job.RenderSheet(1, 72));
        int darkPixels = Enumerable.Range(0, raster.Width).Count(x => raster.GetPixel(x, raster.Height / 2).R < 128);
        Assert.InRange(darkPixels, 1, 3); // The 0.2-point vector becomes two pixels, not an enlarged source pixel.
    }

    [Fact]
    public void PrintOnlySourceFallsBackToRasterContentAndKeepsSelectionOrder() {
        var encryption = new PdfStandardEncryptionOptions("print-user") {
            OwnerPassword = "print-owner", AllowedPermissions = PdfStandardPermissions.Print
        };
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions().SetEncryption(encryption))
            .Paragraph(p => p.AddText("Printable without extraction permission")).ToBytes();
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "print-user" });
        PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest {
            InputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf"),
            PaperSize = PdfPageSizes.A4, Pages = "1", Orientation = PdfPrintOrientation.Landscape
        }, new PdfPrintJobOptions { RasterContentDpi = 150 });
        Assert.Equal(PdfPrintContentKind.Raster, job.ContentKind);
        Assert.Contains(job.Diagnostics, diagnostic => diagnostic.StartsWith("print.raster-content:", StringComparison.Ordinal));
        PdfReadDocument read = PdfReadDocument.Open(ReadContent(job));
        Assert.Single(read.Pages);
        Assert.DoesNotContain("Printable", read.ExtractText());
        PdfPrintSheetImage sheet = job.RenderSheet(1, 72);
        Assert.True(sheet.Width > sheet.Height);
        Assert.Equal(bytes, document.ToBytes());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Reader.Text());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void CompositionRejectsExceededResourceBudgets(int budget) {
        PdfDocument document = PdfDocument.Create(c => {
            c.Page(p => p.Size(OfficeLength.Points(200), OfficeLength.Points(300))); c.Page(p => p.Size(OfficeLength.Points(200), OfficeLength.Points(300)));
        });
        var limits = new PdfPrintJobOptions();
        if (budget == 0) limits.MaximumSheets = 1;
        if (budget == 1) limits.MaximumContentBytes = 64;
        if (budget == 2) limits.MaximumPixelsPerImage = 1;
        Action compose = () => PdfPrintJob.Create(document, new PdfPrintPlanRequest { InputPath = "snapshot.pdf" }, limits);
        if (budget == 1) Assert.Throws<InvalidDataException>(compose);
        else if (budget == 0) Assert.Throws<InvalidOperationException>(compose);
        else {
            PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest { InputPath = "snapshot.pdf" }, limits);
            Assert.Throws<PdfReadLimitException>(() => job.RenderSheet(1, 72));
        }
    }

    [Fact]
    public void CancelledCompositionStopsBeforeRendering() {
        PdfDocument document = PdfDocument.Create(c => c.Page(p => p.Size(OfficeLength.Points(200), OfficeLength.Points(300))));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfPrintJob.Create(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf" }, cancellationToken: cancellation.Token));
        PdfPrintJob job = PdfPrintJob.Create(document, new PdfPrintPlanRequest { InputPath = "snapshot.pdf" });
        Assert.ThrowsAny<OperationCanceledException>(() => job.RenderSheet(1, 72, cancellation.Token));
    }

    private static PdfDocument Square(double size) => PdfDocument.Create(new PdfCreateOptions {
        PageSetup = {
            Size = Paper(size, size),
            Margins = { Left = OfficeLength.Points(0), Right = OfficeLength.Points(0), Top = OfficeLength.Points(0), Bottom = OfficeLength.Points(0) }
        }
    });

    private static PdfPageSize Paper(double width, double height) => new(OfficeLength.Points(width), OfficeLength.Points(height));

    internal static OfficeRasterImage Decode(PdfPrintSheetImage image) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.GetPng(), new OfficeRasterDecodeOptions(), out OfficeRasterImage? raster, out _));
        return raster!;
    }

    private static byte[] ReadContent(PdfPrintJob job) {
        using Stream stream = job.OpenContent();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
