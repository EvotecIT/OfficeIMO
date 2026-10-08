using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRasterNormalizationProjectionContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProjectionSizesOrientedTiffFromItsActualNormalizedPixels(bool applyOrientation) {
        var stored = new OfficeRasterImage(96, 48, OfficeColor.Red);
        stored.SetPixel(0, 0, OfficeColor.Blue);
        byte[] encoded = OfficeTiffCodec.Encode(stored);
        var metadata = OfficeImageMetadata.Read(encoded);
        metadata.SetExifValue(OfficeExifTag.Orientation, (ushort)6);
        byte[] tagged = OfficeImageMetadata.Apply(encoded, metadata);
        OfficeDocumentModel source = Source(tagged);
        Assert.Equal(48, source.Assets[0].Width);
        Assert.Equal(96, source.Assets[0].Height);

        PdfDocumentConversionResult result = source.ToPdfDocumentResult(new PdfProjectionOptions {
            RasterDecodeOptions = new OfficeRasterDecodeOptions { ApplyExifOrientation = applyOrientation }
        });
        OfficeRasterImage expected = applyOrientation ? OfficeRasterTransforms.AutoOrient(stored, 6) : stored;
        AssertArtifactMatchesRaster(result, expected);
        Assert.Equal(48, source.Assets[0].Width);
        Assert.Equal(96, source.Assets[0].Height);
    }

    [Fact]
    public void ProjectionSizesSelectedTiffPageFromItsActualNormalizedPixels() {
        var first = new OfficeRasterImage(64, 128, OfficeColor.Red);
        var selected = new OfficeRasterImage(128, 64, OfficeColor.Blue);
        OfficeDocumentModel source = Source(OfficeTiffCodec.EncodePages(new[] { first, selected }));

        PdfDocumentConversionResult result = source.ToPdfDocumentResult(new PdfProjectionOptions {
            RasterDecodeOptions = new OfficeRasterDecodeOptions { FrameIndex = 1 }
        });

        AssertArtifactMatchesRaster(result, selected);
        Assert.Equal(64, source.Assets[0].Width);
        Assert.Equal(128, source.Assets[0].Height);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProjectionAppliesItsPixelCeilingToTheSelectedTiffPage(bool firstPageIsLarge) {
        var large = new OfficeRasterImage(4001, 2000, OfficeColor.Blue);
        var small = new OfficeRasterImage(96, 48, OfficeColor.Red);
        byte[] encoded = OfficeTiffCodec.EncodePages(
            firstPageIsLarge ? new[] { large, small } : new[] { small, large },
            new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.Deflate });
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { FrameIndex = 1 }, out var coreSelected, out _));
        Assert.NotNull(coreSelected);
        var decodeOptions = new OfficeRasterDecodeOptions { FrameIndex = 1, MaximumDecodedPixels = 50_000_000L };

        PdfDocumentConversionResult result = Source(encoded).ToPdfDocumentResult(
            new PdfProjectionOptions { RasterDecodeOptions = decodeOptions });

        Assert.Equal(50_000_000L, decodeOptions.MaximumDecodedPixels);
        if (firstPageIsLarge) {
            AssertArtifactMatchesRaster(result, small);
        } else {
            Assert.Empty(result.Value.Images.Placements());
            Assert.Contains(result.Warnings, warning => warning.Code == "pdf-projection-asset-listed-not-embedded");
        }
    }

    [Fact]
    public void ProjectionRetainsAStricterCallerPixelCeilingForTheSelectedTiffPage() {
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(16, 16, OfficeColor.Red),
            new OfficeRasterImage(96, 48, OfficeColor.Blue)
        });
        var decodeOptions = new OfficeRasterDecodeOptions { FrameIndex = 1, MaximumDecodedPixels = 4607 };

        PdfDocumentConversionResult result = Source(encoded).ToPdfDocumentResult(
            new PdfProjectionOptions { RasterDecodeOptions = decodeOptions });

        Assert.Empty(result.Value.Images.Placements());
        Assert.Contains(result.Warnings, warning => warning.Code == "pdf-projection-asset-listed-not-embedded");
        Assert.Equal(4607L, decodeOptions.MaximumDecodedPixels);
    }

    private static OfficeDocumentModel Source(byte[] encoded) {
        OfficeImageInfo sourceInfo = OfficeImageReader.Identify(encoded);
        return new OfficeDocumentModel {
            Assets = new[] {
                new OfficeDocumentModelAsset {
                    Id = "raster", Kind = "image", FileName = "raster.tiff", MediaType = "image/tiff",
                    PayloadBytes = encoded, Width = sourceInfo.Width, Height = sourceInfo.Height
                }
            }
        };
    }

    private static void AssertArtifactMatchesRaster(PdfDocumentConversionResult result, OfficeRasterImage expected) {
        PdfImagePlacement placement = Assert.Single(result.Value.Images.Placements());
        PdfExtractedImage embedded = Assert.Single(result.Value.Images.Extract());
        Assert.Equal(expected.Width, embedded.Width);
        Assert.Equal(expected.Height, embedded.Height);
        Assert.Equal(expected.Width * 0.75D, placement.Width, 6);
        Assert.Equal(expected.Height * 0.75D, placement.Height, 6);
        Assert.Equal(expected.GetPixels(), OfficeRasterImageDecoder.Decode(embedded.Bytes).GetPixels());
    }
}
