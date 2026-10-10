using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPictureWorkTests {
    [Fact]
    public void Repeated_picture_decoding_accounts_encoded_payload_work_before_decoding() {
        byte[] png = PublisherPictureWorkFixture.PaddedPng();
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? raster));
        Assert.Equal(1, raster!.Width);
        byte[] input = PublisherPictureWorkFixture.WithImage(PublisherPictureEffectTests.PictureInput(
            new Dictionary<ushort, uint> { [0x013F] = 0x00040004 }), png);
        input = PublisherDrawingFixture.Mutate(input, new Dictionary<ushort, uint> {
            [0x0104] = 1, [0x013F] = 0x00040004
        }, 1, objectId: 367);
        var exception = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxInputBytes = input.Length } }));
        Assert.Contains("image processing byte limit", exception.Message);
    }

    [Fact]
    public void Picture_effect_budget_includes_all_inspected_gif_frames_and_repeated_references() {
        byte[] gif = PublisherPictureWorkFixture.TwoFrameGif();
        Assert.True(OfficeRasterContainerInspector.TryInspect(gif, out OfficeRasterContainerInfo? inventory));
        Assert.Equal(2, inventory!.Count);
        byte[] input = PublisherPictureWorkFixture.WithImage(PublisherPictureEffectTests.PictureInput(
            new Dictionary<ushort, uint> { [0x013F] = 0x00040004 }), gif);
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { MaximumImageProcessingPixels = 3 }));
        PublisherDocument document = PublisherDocument.Load(input, new PublisherReadOptions { MaximumImageProcessingPixels = 4 });
        Assert.Equal(OfficeColor.White, PublisherPictureEffectTests.ProjectedRaster(document, 362).GetPixel(0, 0));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected);
        input = PublisherDrawingFixture.Mutate(input, new Dictionary<ushort, uint> {
            [0x0104] = 1, [0x013F] = 0x00040004
        }, 1, objectId: 367);
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { MaximumImageProcessingPixels = 7 }));
        PublisherDocument.Load(input, new PublisherReadOptions { MaximumImageProcessingPixels = 8 });
    }
}
