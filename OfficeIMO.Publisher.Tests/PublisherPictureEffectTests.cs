using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPictureEffectTests {
    [Theory]
    [InlineData(0x00040004U, 54)]
    [InlineData(0x00020002U, 0)]
    [InlineData(0x00060006U, 0)]
    public void Native_picture_modes_project_pixels_and_preserve_original_assets_and_losses(uint flags, byte expected) {
        byte[] input = PictureInput(new Dictionary<ushort, uint> { [0x013F] = flags });
        byte[][] originals = PublisherDocument.Load(input).Images.Select(image => image.GetBytes()).ToArray();
        var codec = new PictureCodec();
        PublisherDocument document = PublisherDocument.Load(input, new PublisherReadOptions { ImageCodec = codec });
        Assert.All(document.Images.Select((image, index) => (image, index)), entry =>
            Assert.Equal(originals[entry.index], entry.image.GetBytes()));
        Assert.Equal(OfficeColor.FromRgb(expected, expected, expected), ProjectedRaster(document, 362).GetPixel(0, 0));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_PICTURE_EFFECT_APPROXIMATED" && item.Location == "Contents/object/362");
        Assert.Contains(document.ToSvgResult(PicturePageIndex(document, 362)).Report.FidelityDiagnostics,
            item => item.Code == "PUB_PICTURE_EFFECT_APPROXIMATED");
        var pdf = document.ToPdfDocumentResult();
        Assert.Contains(document.ReadReport, pdf.SourceConversionReports);
        Assert.True(pdf.HasLoss);
        Assert.Equal(document.Pages.Count, PdfReadDocument.Open(pdf.ToBytes()).Pages.Count);
        Assert.True(codec.Calls > 0);
    }

    [Fact]
    public void Native_transparent_key_is_applied_before_tone_and_survives_png_projection() {
        PublisherDocument document = PublisherDocument.Load(PictureInput(new Dictionary<ushort, uint> {
            [0x0107] = 0x000000FF, [0x0109] = 32768
        }), new PublisherReadOptions { ImageCodec = new PictureCodec() });
        OfficeRasterImage raster = ProjectedRaster(document, 362);
        Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 0), raster.GetPixel(0, 0));
        Assert.Equal(OfficeColor.White, raster.GetPixel(1, 0));
    }

    [Fact]
    public void Native_gif_effects_use_managed_decoding_and_leave_original_gif_available() {
        byte[] input = PictureInput(new Dictionary<ushort, uint> { [0x013F] = 0x00040004 }, 367);
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal("image/gif", document.Images.Single(image => image.Id == 2).ContentType);
        OfficeRasterImage raster = ProjectedRaster(document, 367);
        Assert.True(raster.Width > 0 && raster.Height > 0);
        foreach (OfficeColor color in new[] { raster.GetPixel(0, 0), raster.GetPixel(raster.Width / 2, raster.Height / 2) }) {
            Assert.Equal(color.R, color.G); Assert.Equal(color.R, color.B);
        }
    }

    [Theory]
    [InlineData(0x0109, 32769U, "PUB_PICTURE_TONE_INVALID")]
    [InlineData(0x0108, 0xFFFFFFFFU, "PUB_PICTURE_TONE_INVALID")]
    [InlineData(0x0107, 0x10000001U, "PUB_PICTURE_TRANSPARENT_COLOR_UNRESOLVED")]
    [InlineData(0x011A, 0x00030201U, "PUB_PICTURE_RECOLOR_UNASSESSED")]
    public void Native_invalid_or_unqualified_controls_report_loss_without_disabling_valid_modes(ushort property, uint value, string diagnostic) {
        PublisherDocument document = PublisherDocument.Load(PictureInput(new Dictionary<ushort, uint> {
            [property] = value, [0x013F] = 0x00040004
        }), new PublisherReadOptions { ImageCodec = new PictureCodec() });
        Assert.Equal(OfficeColor.FromRgb(54, 54, 54), ProjectedRaster(document, 362).GetPixel(0, 0));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == diagnostic);
    }

    [Theory]
    [InlineData(0x0117, 0x20000001U, true)]
    [InlineData(0x011D, 0x20000001U, true)]
    [InlineData(0x0115, 0xFFFFFFFFU, false)]
    [InlineData(0x011B, 0xFFFFFFFFU, false)]
    [InlineData(0x0117, 0x20000000U, false)]
    [InlineData(0x011D, 0x20000000U, false)]
    [InlineData(0x0116, 0xFFFFFFFFU, false)]
    [InlineData(0x011C, 0xFFFFFFFFU, false)]
    public void Native_extended_controls_ignore_reserved_and_inactive_values(ushort property, uint value, bool reported) {
        PublisherDocument document = PublisherDocument.Load(PictureInput(new Dictionary<ushort, uint> {
            [property] = value, [0x013F] = 0x00040004
        }), new PublisherReadOptions { ImageCodec = new PictureCodec() });
        Assert.Equal(reported, document.ReadReport.FidelityDiagnostics.Any(item => item.Code == "PUB_PICTURE_RECOLOR_UNASSESSED"
            && item.Location == "Contents/object/362"));
        Assert.Equal(OfficeColor.FromRgb(54, 54, 54), ProjectedRaster(document, 362).GetPixel(0, 0));
    }

    [Fact]
    public void Picture_effect_pixels_obey_per_image_and_cumulative_bounds() {
        byte[] input = PictureInput(new Dictionary<ushort, uint> { [0x013F] = 0x00040004 });
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { ImageCodec = new PictureCodec(), MaximumRasterPixels = 1 }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { ImageCodec = new PictureCodec(), MaximumImageProcessingPixels = 3 }));
        PublisherDocument document = PublisherDocument.Load(input,
            new PublisherReadOptions { ImageCodec = new PictureCodec(), MaximumImageProcessingPixels = 4 });
        Assert.Equal(2, ProjectedRaster(document, 362).Width);
        byte[] managed = PictureInput(new Dictionary<ushort, uint> { [0x013F] = 0x00040004 }, 367);
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(managed, new PublisherReadOptions { MaximumRasterPixels = 1 }));
    }

    [Fact]
    public void Picture_effect_codec_cancellation_rejects_the_operation() {
        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => PublisherDocument.Load(
            PictureInput(new Dictionary<ushort, uint> { [0x013F] = 0x00040004 }),
            new PublisherReadOptions { ImageCodec = new PictureCodec(source) }, source.Token));
    }

    [Fact]
    public void Repeated_picture_references_share_the_cumulative_processing_bound() {
        byte[] input = PictureInput(new Dictionary<ushort, uint> { [0x013F] = 0x00040004 });
        input = PublisherDrawingFixture.Mutate(input, new Dictionary<ushort, uint> {
            [0x0104] = 1, [0x013F] = 0x00040004
        }, 1, objectId: 367);
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { ImageCodec = new PictureCodec(), MaximumImageProcessingPixels = 7 }));
        PublisherDocument document = PublisherDocument.Load(input,
            new PublisherReadOptions { ImageCodec = new PictureCodec(), MaximumImageProcessingPixels = 8 });
        Assert.Equal(OfficeColor.FromRgb(54, 54, 54), ProjectedRaster(document, 362).GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgb(54, 54, 54), ProjectedRaster(document, 367).GetPixel(0, 0));
    }

    internal static byte[] PictureInput(Dictionary<ushort, uint> values, uint objectId = 362) =>
        PublisherDrawingFixture.Mutate(values, 1, fixture: "SampleBrochure.pub", objectId: objectId);

    internal static OfficeRasterImage ProjectedRaster(PublisherDocument document, uint objectId) {
        OfficeDrawingImage image = PublisherNativeTests.Elements(document.Pages[PicturePageIndex(document, objectId)].Drawing)
            .OfType<OfficeDrawingImage>().Single(item => item.SourceElementIds?.Contains("publisher-object-" + objectId) == true);
        Assert.Equal("image/png", image.ContentType);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        return raster!;
    }

    internal static int PicturePageIndex(PublisherDocument document, uint objectId) => document.Pages
        .Select((page, index) => (page, index)).Single(entry => PublisherNativeTests.Elements(entry.page.Drawing)
            .OfType<OfficeDrawingImage>().Any(image => image.SourceElementIds?.Contains("publisher-object-" + objectId) == true)).index;

    private sealed class PictureCodec : IOfficeRasterImageCodec {
        private readonly CancellationTokenSource? _cancel;
        internal PictureCodec(CancellationTokenSource? cancel = null) => _cancel = cancel;
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++; encodedBytes[0] = 0; _cancel?.Cancel();
            image = new OfficeRasterImage(2, 1, OfficeColor.FromRgb(255, 0, 0));
            image.SetPixel(1, 0, OfficeColor.FromRgb(0, 255, 0));
            return true;
        }
    }
}
