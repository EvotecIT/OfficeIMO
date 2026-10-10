using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGeometryGuideTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_guides_resolve_adjustments_and_previous_results_without_losing_following_complex_properties(bool excludeHeaders) {
        byte[] input = Input(new[] { ((ushort)0x2000, (ushort)0x0147, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0) }, excludeHeaders: excludeHeaders);
        PublisherDocument publication = PublisherDocument.Load(input);
        OfficeShape shape = Artwork(publication);
        Assert.Equal(OfficeShapeKind.Path, shape.Kind);
        Assert.Equal(new OfficePoint(shape.Width / 4, shape.Height / 4), shape.PathCommands[0].Point);
        Assert.Equal(new OfficePoint(shape.Width * 3 / 4, shape.Height / 4), shape.PathCommands[1].Point);
        Assert.DoesNotContain(publication.ReadReport.FidelityDiagnostics, item => item.Code.StartsWith("PUB_CUSTOM_PATH_GUIDE"));
        var svg = publication.ToSvgResult();
        Assert.Contains(svg.Report.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_RENDERING_UNQUALIFIED"
            && item.Message.Contains("producer rounding") && item.Location == "Contents/object/293");
        Assert.Throws<OfficeConversionException>(() => svg.RequireNoLoss());
        var pdf = publication.ToPdfDocumentResult();
        Assert.Contains(publication.ReadReport, pdf.SourceConversionReports);
        Assert.Equal(publication.Pages.Count, PdfReadDocument.Open(pdf.ToBytes()).Pages.Count);
        Assert.NotEmpty(publication.TextStories);
    }

    [Theory]
    [InlineData(0x2000, 0x0400, "PUB_CUSTOM_PATH_GUIDES_INVALID")]
    [InlineData(0x0011, 10, "PUB_CUSTOM_PATH_GUIDE_FORMULA_UNSUPPORTED")]
    [InlineData(0x2000, 0x04F8, "PUB_CUSTOM_PATH_GUIDE_PARAMETER_UNSUPPORTED")]
    [InlineData(0x0001, 10, "PUB_CUSTOM_PATH_GUIDES_INVALID")]
    public void Native_invalid_or_unsupported_guides_preserve_text_and_report_the_geometry_fallback(
        ushort operation, ushort first, string code) {
        var publication = PublisherDocument.Load(Input(new[] { (operation, first, (ushort)0, (ushort)0) }));
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.NotEmpty(publication.TextStories);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == code
            && item.Location == "Contents/object/293" && item.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Native_guides_retain_recovery_limits_in_the_complete_reader() {
        var records = Enumerable.Range(0, 128).Select(index => ((ushort)0, (ushort)index, (ushort)0, (ushort)0)).ToArray();
        var error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(Input(records),
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 100 } }));
        Assert.Contains("custom path item work limit", error.Message);
    }

    [Fact]
    public void Limousine_scaling_keeps_an_explicit_fallback_even_for_literal_paths() {
        byte[] input = PublisherDrawingFixture.Mutate(new() { [0x0153] = 100, [0x0144] = 1 }, 0,
            new() { [0x0145] = PublisherCustomPathTests.Vertices(new[] { (0, 0), (21600, 0), (0, 21600) }) });
        var publication = PublisherDocument.Load(input);
        Assert.Equal(OfficeShapeKind.Rectangle, Artwork(publication).Kind);
        Assert.Contains(publication.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_SCALING_UNSUPPORTED");
    }

    [Fact]
    public void Guided_picture_masks_use_the_same_geometry_and_keep_original_assets() {
        byte[] original = PublisherPictureEffectTests.PictureInput(new());
        var values = new Dictionary<ushort, uint> { [0x140] = 0, [0x141] = 0, [0x142] = 100, [0x143] = 100,
            [0x144] = 1, [0x147] = 25 };
        byte[] input = PublisherDrawingFixture.Mutate(original, values, 0, Complex(new[] {
            ((ushort)0x2000, (ushort)0x0147, (ushort)0, (ushort)0),
            ((ushort)0x2000, (ushort)0x0400, (ushort)50, (ushort)0) }), 362);
        var options = new PublisherReadOptions { ImageCodec = new SolidCodec() };
        PublisherDocument publication = PublisherDocument.Load(input, options);
        PublisherDocument control = PublisherDocument.Load(original, options);
        Assert.Equal(control.Images.SelectMany(image => image.GetBytes()), publication.Images.SelectMany(image => image.GetBytes()));
        var group = publication.Pages.SelectMany(page => PublisherNativeTests.Elements(page.Drawing))
            .OfType<OfficeDrawingGroup>().Single(item => item.SourceElementIds?.Contains("publisher-object-362") == true);
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(new OfficePoint(group.ClipPath.Width / 4, group.ClipPath.Height / 4), group.ClipPath.Commands[0].Point);
        var drawing = new OfficeDrawing(group.ClipPath.Width, group.ClipPath.Height)
            .AddClippedDrawing(group.InnerDrawing, 0, 0, group.ClipPath);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel((int)(raster.Width * 0.3), (int)(raster.Height * 0.3)));
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.8), (int)(raster.Height * 0.8)).A);
        Assert.Equal(0, raster.GetPixel((int)(raster.Width * 0.1), (int)(raster.Height * 0.1)).A);
    }

    internal static byte[] Input((ushort Operation, ushort First, ushort Second, ushort Third)[] records,
        bool excludeHeaders = false) => PublisherDrawingFixture.Mutate(new() {
            [0x140] = 0, [0x141] = 0, [0x142] = 100, [0x143] = 100, [0x144] = 1,
            [0x147] = 25, [0x181] = 0x000000FF, [0x1BF] = 0x00100010
        }, 0, Complex(records), arrayLengthsExcludeHeader: excludeHeaders);

    private static Dictionary<ushort, byte[]> Complex((ushort Operation, ushort First, ushort Second, ushort Third)[] records) => new() {
        [0x0145] = PublisherCustomPathTests.Vertices(new[] { (Guide(0), Guide(0)), (Guide(1), Guide(0)), (Guide(0), Guide(1)) }),
        [0x0156] = GuideArray(records),
        [0x0380] = System.Text.Encoding.Unicode.GetBytes("Guided artwork\0")
    };
    internal static byte[] GuideArray((ushort Operation, ushort First, ushort Second, ushort Third)[] records) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(checked((ushort)records.Length)); writer.Write(checked((ushort)records.Length)); writer.Write((ushort)8);
        foreach (var record in records) {
            writer.Write(record.Operation); writer.Write(record.First); writer.Write(record.Second); writer.Write(record.Third);
        }
        return stream.ToArray();
    }
    private static int Guide(int index) => unchecked((int)0x80000000) | index;
    private static OfficeShape Artwork(PublisherDocument publication) => PublisherNativeTests.Elements(publication.Pages[0].Drawing)
        .OfType<OfficeDrawingShape>().Single(item => item.SourceElementIds?.Contains("publisher-object-293") == true).Shape;
    private sealed class SolidCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = new OfficeRasterImage(20, 20, OfficeColor.Red); return true;
        }
    }
}
