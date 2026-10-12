using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageDisplayRendererTests {
    private static readonly PdfPageRegionLimits Limits = new();

    private static PdfDocument CreateDocument(PdfOptions? options = null) => PdfDocument.Load(
        PdfDocument.Create(options ?? new PdfOptions { PageWidth = 300, PageHeight = 200, DefaultFontSize = 11 })
            .Paragraph(paragraph => paragraph.Text("Tiles of one page join without seams at any scale."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            })
            .ToBytes());

    private static (int Width, int Height) PageImageSize(PdfDocument document, double scale) {
        OfficeDrawing drawing = document.Render.Drawing(1);
        return ((int)Math.Ceiling(drawing.Width * scale), (int)Math.Ceiling(drawing.Height * scale));
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(1.7D)]
    [InlineData(3.3D)]
    public void TilesJoinIntoTheWholePage(double scale) {
        PdfDocument document = CreateDocument();
        var renderer = new PdfPageDisplayRenderer(document, pageCapacity: 2, imageCapacityBytes: 64L * 1024 * 1024);
        (int width, int height) = PageImageSize(document, scale);
        byte[] whole = renderer.RenderRegion(1, scale, 0, 0, width, height, Limits, default).GetPixels();

        int splitX = width / 2, splitY = height / 3;
        var assembled = new byte[whole.Length];
        foreach ((int x, int y, int w, int h) in new[] {
                     (0, 0, splitX, splitY), (splitX, 0, width - splitX, splitY),
                     (0, splitY, splitX, height - splitY), (splitX, splitY, width - splitX, height - splitY) }) {
            byte[] tile = renderer.RenderRegion(1, scale, x, y, w, h, Limits, default).GetPixels();
            for (int row = 0; row < h; row++) {
                Buffer.BlockCopy(tile, row * w * 4, assembled, ((y + row) * width + x) * 4, w * 4);
            }
        }

        Assert.Equal(whole, assembled);
    }

    [Fact]
    public void TheWholePageIsTheRenderOfItsDrawingOnWhite() {
        PdfDocument document = CreateDocument();
        var renderer = new PdfPageDisplayRenderer(document, pageCapacity: 1, imageCapacityBytes: 16L * 1024 * 1024);
        (int width, int height) = PageImageSize(document, 2D);

        OfficeRasterImage region = renderer.RenderRegion(1, 2D, 0, 0, width, height, Limits, default);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(document.Render.Drawing(1), 2D, OfficeColor.White);

        Assert.Equal((expected.Width, expected.Height), (region.Width, region.Height));
        Assert.Equal(expected.GetPixels(), region.GetPixels());
    }

    [Fact]
    public void RepeatedRendersReuseTheDecodedImages() {
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "openpreserve-high-resolution-image.pdf");
        PdfDocument document = PdfDocument.Load(File.ReadAllBytes(path));
        var renderer = new PdfPageDisplayRenderer(document, pageCapacity: 1, imageCapacityBytes: 64L * 1024 * 1024);
        (int width, int height) = PageImageSize(document, 1.5D);

        byte[] first = renderer.RenderRegion(1, 1.5D, 0, 0, width, height, Limits, default).GetPixels();
        long cached = renderer.CachedImageBytes;
        byte[] second = renderer.RenderRegion(1, 1.5D, 0, 0, width, height, Limits, default).GetPixels();

        Assert.True(cached > 0L);
        Assert.Equal(cached, renderer.CachedImageBytes);
        Assert.Equal(first, second);
    }

    [Fact]
    public void DocumentsThatRestrictCopyingRenderAsPixels() {
        var encryption = new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None };
        byte[] restricted = CreateDocument(new PdfOptions { PageWidth = 300, PageHeight = 200 }.SetEncryption(encryption)).ToBytes();
        PdfDocument document = PdfDocument.Load(restricted, new PdfLoadOptions { Password = "open" });
        var renderer = new PdfPageDisplayRenderer(document, pageCapacity: 1, imageCapacityBytes: 16L * 1024 * 1024);

        OfficeRasterImage image = renderer.RenderRegion(1, 1D, 0, 0, 300, 200, Limits, default);

        Assert.Contains(image.GetPixels().Select((value, index) => (value, index)), pixel => pixel.index % 4 != 3 && pixel.value < 128);
    }

    [Fact]
    public void RegionsOutsideThePageOrBeyondTheLimitAreRefused() {
        PdfDocument document = CreateDocument();
        var renderer = new PdfPageDisplayRenderer(document, pageCapacity: 1, imageCapacityBytes: 16L * 1024 * 1024);

        Assert.Throws<ArgumentOutOfRangeException>(() => renderer.RenderRegion(1, 1D, 200, 0, 200, 100, Limits, default));
        Assert.Throws<ArgumentOutOfRangeException>(() => renderer.RenderRegion(99, 1D, 0, 0, 10, 10, Limits, default));
        PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() => renderer.RenderRegion(1, 1D, 0, 0, 300, 200,
            new PdfPageRegionLimits { MaximumRegionPixels = 1_000 }, default));
        Assert.Equal(PdfReadLimitKind.RenderPixels, limit.Kind);
    }
}
