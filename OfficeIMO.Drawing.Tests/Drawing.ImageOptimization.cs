using OfficeIMO.Drawing;
using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingImageOptimizationTests {
        [Fact]
        public void OfficeRasterResampler_NearestNeighborPreservesSourceQuadrants() {
            var source = new OfficeRasterImage(2, 2);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Lime);
            source.SetPixel(0, 1, OfficeColor.Blue);
            source.SetPixel(1, 1, OfficeColor.White);

            OfficeRasterImage resized = OfficeRasterResampler.Resize(source, 4, 4, OfficeRasterResamplingMode.NearestNeighbor);

            Assert.Equal(OfficeColor.Red, resized.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Lime, resized.GetPixel(3, 0));
            Assert.Equal(OfficeColor.Blue, resized.GetPixel(0, 3));
            Assert.Equal(OfficeColor.White, resized.GetPixel(3, 3));
        }

        [Fact]
        public void OfficeRasterResampler_BilinearUsesPremultipliedAlpha() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.FromRgba(0, 0, 255, 0));
            source.SetPixel(1, 0, OfficeColor.Red);

            OfficeColor blended = OfficeRasterResampler.Resize(source, 1, 1).GetPixel(0, 0);

            Assert.InRange(blended.R, 254, 255);
            Assert.InRange(blended.G, 0, 1);
            Assert.InRange(blended.B, 0, 1);
            Assert.InRange(blended.A, 127, 128);
        }

        [Fact]
        public void OfficeRasterResampler_BilinearClampsEdgeCoordinatesBeforeWeighting() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Blue);

            OfficeRasterImage resized = OfficeRasterResampler.Resize(source, 4, 1);

            Assert.Equal(OfficeColor.Red, resized.GetPixel(0, 0));
            Assert.Equal(OfficeColor.Blue, resized.GetPixel(3, 0));
        }

        [Fact]
        public void OfficeJpegCodec_RoundTripsDimensionsAndRepresentativeColors() {
            OfficeRasterImage source = CreateQuadrantImage(32, 24);

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 92,
                Subsampling = OfficeJpegSubsampling.Y444
            });

            Assert.True(OfficeJpegCodec.IsJpeg(jpeg));
            Assert.True(OfficeRasterImageDecoder.TryDecode(jpeg, out OfficeRasterImage? decoded));
            Assert.NotNull(decoded);
            Assert.Equal(32, decoded!.Width);
            Assert.Equal(24, decoded.Height);
            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.Red, 20);
            AssertColorNear(decoded.GetPixel(27, 4), OfficeColor.Lime, 20);
            AssertColorNear(decoded.GetPixel(4, 19), OfficeColor.Blue, 20);
            AssertColorNear(decoded.GetPixel(27, 19), OfficeColor.White, 20);
        }

        [Theory]
        [InlineData(OfficeJpegSubsampling.Y444, false, false)]
        [InlineData(OfficeJpegSubsampling.Y422, false, true)]
        [InlineData(OfficeJpegSubsampling.Y420, true, false)]
        [InlineData(OfficeJpegSubsampling.Y420, true, true)]
        public void OfficeJpegCodec_EncodesManagedVariants(OfficeJpegSubsampling subsampling, bool progressive, bool optimizeHuffman) {
            OfficeRasterImage source = CreateQuadrantImage(37, 29);

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 88,
                Subsampling = subsampling,
                Progressive = progressive,
                OptimizeHuffman = optimizeHuffman
            });
            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg, new OfficeJpegDecodeOptions(highQualityChroma: true));

            Assert.Equal(source.Width, decoded.Width);
            Assert.Equal(source.Height, decoded.Height);
            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.Red, 28);
            AssertColorNear(decoded.GetPixel(32, 24), OfficeColor.White, 28);
        }

        [Fact]
        public void OfficeJpegCodec_ProgressiveColorUsesNonInterleavedAcScans() {
            byte[] jpeg = OfficeJpegCodec.Encode(CreateQuadrantImage(37, 29), new OfficeJpegEncodeOptions {
                Progressive = true,
                Subsampling = OfficeJpegSubsampling.Y420
            });

            var scans = ReadStartOfScanHeaders(jpeg);

            Assert.Equal(4, scans.Count);
            Assert.Equal((3, 0), scans[0]);
            Assert.All(scans.Skip(1), scan => {
                Assert.Equal(1, scan.ComponentCount);
                Assert.Equal(1, scan.SpectralStart);
            });
        }

        [Fact]
        public void OfficeJpegCodec_RejectsDimensionsBeyondJpegHeaderLimits() {
            var source = new OfficeRasterImage(ushort.MaxValue + 1, 1, OfficeColor.Red);

            ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => OfficeJpegCodec.Encode(source));

            Assert.Equal("width", exception.ParamName);
            Assert.Contains("65535", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OfficeJpegCodec_RejectsExcessiveSamplingBeforeDecoderStateAllocation() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(1, 1, OfficeColor.Red),
                new OfficeJpegEncodeOptions { Subsampling = OfficeJpegSubsampling.Y444 });
            int startOfFrame = FindMarker(jpeg, 0xC0);
            Assert.True(startOfFrame > 0);
            jpeg[startOfFrame + 11] = 0x44;

            FormatException exception = Assert.Throws<FormatException>(() => OfficeJpegCodec.Decode(jpeg));

            Assert.Contains("sampling", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OfficeJpegCodec_RejectsDuplicateScanComponentsBeforeDecoding() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(1, 1, OfficeColor.Red),
                new OfficeJpegEncodeOptions { Subsampling = OfficeJpegSubsampling.Y444 });
            int startOfScan = FindMarker(jpeg, 0xDA);
            Assert.True(startOfScan > 0);
            jpeg[startOfScan + 7] = jpeg[startOfScan + 5];

            FormatException exception = Assert.Throws<FormatException>(() => OfficeJpegCodec.Decode(jpeg));

            Assert.Contains("Duplicate", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OfficeJpegCodec_DecodesBaselineComponentsStoredInSeparateScans() {
            byte[] jpeg = BuildSeparateComponentBaselineJpeg();

            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);
            OfficeColor pixel = decoded.GetPixel(0, 0);

            Assert.Equal(1, decoded.Width);
            Assert.Equal(1, decoded.Height);
            Assert.InRange(pixel.R, 127, 129);
            Assert.InRange(pixel.G, 127, 129);
            Assert.InRange(pixel.B, 127, 129);
        }

        [Fact]
        public void OfficeJpegMetadata_SnapshotsCallerBuffers() {
            byte[] exif = { 1, 2, 3 };
            var metadata = new OfficeJpegMetadata(exif: exif);

            exif[0] = 9;
            byte[] exposed = metadata.Exif!;
            exposed[1] = 8;

            Assert.Equal(new byte[] { 1, 2, 3 }, metadata.Exif);
        }

        [Fact]
        public void OfficeImageOptimizerPreservesSelectedJpegMetadataAndNormalizesOrientation() {
            byte[] xmp = System.Text.Encoding.UTF8.GetBytes(
                "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"/></x:xmpmeta>");
            byte[] icc = CreateMinimalOptimizationIccProfile();
            var source = new OfficeRasterImage(2, 1, OfficeColor.CornflowerBlue);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Metadata = new OfficeJpegMetadata(CreateExifOrientation(6), xmp, icc),
                DpiX = 144,
                DpiY = 72
            });
            var request = new OfficeImageOptimizationRequest(1, 1) {
                OutputFormat = OfficeImageFormat.Jpeg,
                KeepOriginalWhenNotSmaller = false,
                MetadataPolicy = OfficeImageMetadataPolicy.SelectiveCopy,
                MetadataSelection = OfficeImageMetadataKinds.Exif | OfficeImageMetadataKinds.Xmp |
                                    OfficeImageMetadataKinds.Icc | OfficeImageMetadataKinds.Orientation |
                                    OfficeImageMetadataKinds.Resolution
            };

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(jpeg, request);

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.False(result.Metadata.HasLoss);
            Assert.Equal(request.MetadataSelection, result.Metadata.Preserved);
            Assert.Equal(OfficeImageMetadataKinds.Orientation, result.Metadata.Normalized);
            Assert.True(OfficeImageOrientationNormalizer.TryRead(result.Bytes, out OfficeImageOrientation orientation));
            Assert.Equal(OfficeImageOrientation.Normal, orientation);
            Assert.True(OfficeImageReader.TryValidateContent(result.Bytes, "optimized.jpg", out _));
            Assert.True(ContainsSequence(result.Bytes, CreateExifOrientation(1)));
            Assert.True(ContainsSequence(result.Bytes, xmp));
            Assert.True(ContainsSequence(result.Bytes, icc));
        }

        [Fact]
        public void OfficeImageOptimizerReportsNonRgbIccAsLostWhenWritingRgbJpeg() {
            byte[] grayIcc = CreateMinimalOptimizationIccProfile("GRAY");
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(icc: grayIcc)
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageMetadataKinds.Icc,
                result.Metadata.Source & OfficeImageMetadataKinds.Icc);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Icc);
            Assert.Equal(OfficeImageMetadataKinds.Icc,
                result.Metadata.Lost & OfficeImageMetadataKinds.Icc);
            Assert.Equal(OfficeImageMetadataKinds.None,
                OfficeImageMetadataInspector.Inspect(result.Bytes, OfficeImageFormat.Jpeg).Kinds &
                OfficeImageMetadataKinds.Icc);
        }

        [Fact]
        public void OfficeImageOptimizerReportsNormalizedTiffOrientationAsPreserved() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Blue);
            byte[] tiff = OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                DpiX = 300D,
                DpiY = 150D
            });
            int orientationEntry = FindClassicTiffEntry(tiff, 274);
            WriteLittleEndian(tiff, orientationEntry + 8, 6);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                tiff,
                new OfficeImageOptimizationRequest(1, 2) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal((1, 2), (result.Final.Width, result.Final.Height));
            Assert.Equal(OfficeImageMetadataKinds.Orientation,
                result.Metadata.Normalized & OfficeImageMetadataKinds.Orientation);
            Assert.Equal(OfficeImageMetadataKinds.Orientation,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Orientation);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Lost & OfficeImageMetadataKinds.Orientation);
        }

        [Fact]
        public void OfficeImageOptimizerReportsStrippedAndUnsupportedMetadata() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(2, 2, OfficeColor.Red),
                new OfficeJpegEncodeOptions { Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(3)) });
            var strip = new OfficeImageOptimizationRequest(1, 1) {
                OutputFormat = OfficeImageFormat.Jpeg,
                KeepOriginalWhenNotSmaller = false,
                MetadataPolicy = OfficeImageMetadataPolicy.Strip
            };

            OfficeImageOptimizationResult stripped = OfficeImageOptimizer.Optimize(jpeg, strip);

            Assert.True(stripped.Metadata.PolicyApplied);
            Assert.Equal(OfficeImageMetadataKinds.None, stripped.Metadata.Requested);
            Assert.NotEqual(OfficeImageMetadataKinds.None, stripped.Metadata.Stripped);
            Assert.False(stripped.Metadata.HasLoss);
        }

        [Fact]
        public void OfficeImageOptimizerOmitsJfifDensityWhenSelectiveCopyExcludesResolution() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions { DpiX = 144D, DpiY = 72D });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.SelectiveCopy,
                    MetadataSelection = OfficeImageMetadataKinds.None
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.False(ContainsSequence(result.Bytes, new byte[] {
                (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0
            }));
            Assert.Equal(OfficeImageMetadataKinds.None,
                OfficeImageMetadataInspector.Inspect(result.Bytes, OfficeImageFormat.Jpeg).Kinds &
                OfficeImageMetadataKinds.Resolution);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Png)]
        [InlineData(OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageFormat.Webp)]
        public void OfficeImageOptimizerOmitsResolutionMetadataFromManagedOutputsWhenStripped(
            OfficeImageFormat outputFormat) {
            byte[] source = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.SteelBlue));

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                source,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = outputFormat,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Strip
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageMetadataKinds.None,
                OfficeImageMetadataInspector.Inspect(result.Bytes, outputFormat).Kinds &
                OfficeImageMetadataKinds.Resolution);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Png)]
        [InlineData(OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageFormat.Webp)]
        public void ExplicitOutputResolutionOverridesMetadataStripping(OfficeImageFormat outputFormat) {
            byte[] source = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.SteelBlue));

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                source,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = outputFormat,
                    OutputDpiX = 300D,
                    OutputDpiY = 150D,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Strip
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                OfficeImageMetadataInspector.Inspect(result.Bytes, outputFormat).Kinds &
                OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void ExplicitJpegResolutionRewritesCopiedExifDensity() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: CreateExifWithResolution(72, 96)),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    OutputDpiX = 300D,
                    OutputDpiY = 150D,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.NotNull(output.Exif);
            byte[] exifOnly = OfficeJpegCodec.Encode(
                new OfficeRasterImage(1, 1, OfficeColor.White),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: output.Exif),
                    WriteJfifHeader = false
                });
            OfficeImageMetadataSnapshot copiedExif = OfficeImageMetadataInspector.Inspect(
                exifOnly,
                OfficeImageFormat.Jpeg);

            Assert.InRange(copiedExif.PhysicalDpiX!.Value, 299.99D, 300.01D);
            Assert.InRange(copiedExif.PhysicalDpiY!.Value, 149.99D, 150.01D);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Normalized & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void AxisSwappingJpegOrientationRewritesCopiedExifDensity() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 2, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(
                        exif: CreateExifWithOrientationAndResolution(6, 300, 150)),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(1, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);

            Assert.True(OfficeImageOrientationNormalizer.TryRead(
                result.Bytes, out OfficeImageOrientation outputOrientation));
            Assert.Equal(OfficeImageOrientation.Normal, outputOrientation);
            Assert.InRange(output.PhysicalDpiX!.Value, 149.99D, 150.01D);
            Assert.InRange(output.PhysicalDpiY!.Value, 299.99D, 300.01D);
            Assert.InRange(result.Final.DpiX, 149.99D, 150.01D);
            Assert.InRange(result.Final.DpiY, 299.99D, 300.01D);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Normalized & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void ExplicitJpegResolutionDropsExifWithAliasedDensityStorage() {
            byte[] exif = CreateExifWithResolution(72, 96);
            WriteLittleEndianUInt32(exif, 30, 50);
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: exif),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    OutputDpiX = 300D,
                    OutputDpiY = 150D,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.Null(output.Exif);
            Assert.InRange(result.Final.DpiX, 299.99D, 300.01D);
            Assert.InRange(result.Final.DpiY, 149.99D, 150.01D);
            Assert.Equal(OfficeImageMetadataKinds.Exif,
                result.Metadata.Lost & OfficeImageMetadataKinds.Exif);
        }

        [Fact]
        public void ExplicitJpegResolutionDropsExifWhenDensityAliasesAnotherTiffValue() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: CreateExifWithResolutionAliasedToWhitePoint()),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    OutputDpiX = 300D,
                    OutputDpiY = 150D,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.Null(output.Exif);
            Assert.Equal(OfficeImageMetadataKinds.Exif,
                result.Metadata.Lost & OfficeImageMetadataKinds.Exif);
        }

        [Fact]
        public void ExifDensityRewriteRejectsAliasWithLongSubIfdTable() {
            byte[] exif = CreateExifWithResolutionAliasedToLongSubIfd();

            Assert.True(OfficeTiffStructureValidator.TryValidateExif(exif, 0, exif.Length));
            Assert.False(OfficeExifMetadataEditor.TryRewritePhysicalResolution(
                exif, 300D, 150D, out _));
        }

        [Fact]
        public void JpegOptimizationDropsExifWhenOrientationAliasesAnotherTiffValue() {
            byte[] exif = CreateExifWithOrientationAliasedToWhitePoint();
            Assert.False(OfficeImageOrientationNormalizer.TryNeutralizeExifOrientation(exif, out _));
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(2, 1, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: exif),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.Null(output.Exif);
            Assert.Equal(OfficeImageMetadataKinds.Exif,
                result.Metadata.Lost & OfficeImageMetadataKinds.Exif);
        }

        [Fact]
        public void JpegOptimizationDropsExifWithDuplicateOrientationEntries() {
            byte[] exif = CreateExifWithDuplicateOrientations();
            Assert.False(OfficeImageOrientationNormalizer.TryNeutralizeExifOrientation(exif, out _));
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(2, 1, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: exif),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.Null(output.Exif);
            Assert.Equal(OfficeImageMetadataKinds.Exif,
                result.Metadata.Lost & OfficeImageMetadataKinds.Exif);
        }

        [Fact]
        public void ExplicitFractionalJpegResolutionUsesOneRepresentableDensity() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: CreateExifWithResolution(72, 96)),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    OutputDpiX = 300.25D,
                    OutputDpiY = 150.25D,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.NotNull(output.Exif);
            byte[] exifOnly = OfficeJpegCodec.Encode(
                new OfficeRasterImage(1, 1, OfficeColor.White),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: output.Exif),
                    WriteJfifHeader = false
                });
            OfficeImageMetadataSnapshot copiedExif = OfficeImageMetadataInspector.Inspect(
                exifOnly,
                OfficeImageFormat.Jpeg);

            Assert.Equal(300D, result.Final.DpiX);
            Assert.Equal(150D, result.Final.DpiY);
            Assert.Equal(300D, copiedExif.PhysicalDpiX);
            Assert.Equal(150D, copiedExif.PhysicalDpiY);
        }

        [Fact]
        public void SelectiveResolutionStrippingDoesNotCopyResolutionBearingExif() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: CreateExifWithResolution())
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.SelectiveCopy,
                    MetadataSelection = OfficeImageMetadataKinds.Exif
                });

            OfficeImageMetadataSnapshot output = OfficeImageMetadataInspector.Inspect(
                result.Bytes,
                OfficeImageFormat.Jpeg);
            Assert.Equal(OfficeImageMetadataKinds.None, output.Kinds & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.Exif, result.Metadata.Lost & OfficeImageMetadataKinds.Exif);
        }

        [Fact]
        public void OfficeImageOptimizerPreservesExifOnlyPhysicalResolutionAcrossFormats() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: CreateExifWithResolution(300, 150)),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.InRange(result.Final.DpiX, 299.9D, 300.1D);
            Assert.InRange(result.Final.DpiY, 149.9D, 150.1D);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Lost & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void OfficeImageOptimizerReportsIncompletePhysicalExifResolutionAsLost() {
            byte[] exif = CreateExifWithResolution(300, 150);
            WriteLittleEndianUInt32(exif, 30, uint.MaxValue);
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(exif: exif),
                    WriteJfifHeader = false
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.InRange(result.Final.DpiX, 95.9D, 96.1D);
            Assert.InRange(result.Final.DpiY, 95.9D, 96.1D);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Lost & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void OfficeImageOptimizerReportsExtendedJpegXmpAsLossWhenReencoding() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Metadata = new OfficeJpegMetadata(xmp: System.Text.Encoding.UTF8.GetBytes("<x:xmpmeta />"))
                });
            byte[] extended = InsertApp1SegmentAfterStartOfImage(
                jpeg,
                System.Text.Encoding.ASCII.GetBytes(
                    "http://ns.adobe.com/xmp/extension/\0" +
                    "0123456789ABCDEF0123456789ABCDEF" +
                    "\0\0\0\x01\0\0\0\0X"));

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                extended,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageMetadataKinds.Xmp, result.Metadata.Lost & OfficeImageMetadataKinds.Xmp);
            Assert.False(OfficeImageMetadataInspector.Inspect(result.Bytes, OfficeImageFormat.Jpeg).HasExtendedJpegXmp);
        }

        [Fact]
        public void OfficeImageOptimizerReportsJpegMetadataFoundBetweenProgressiveScans() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions { Progressive = true });
            byte[] withComment = InsertJpegSegmentBeforeSecondScan(
                jpeg,
                marker: 0xFE,
                System.Text.Encoding.ASCII.GetBytes("between scans"));

            Assert.True(OfficeJpegCodec.TryDecode(withComment, out _));
            Assert.Equal(
                OfficeImageMetadataKinds.Comments,
                OfficeImageMetadataInspector.Inspect(withComment, OfficeImageFormat.Jpeg).Kinds &
                OfficeImageMetadataKinds.Comments);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                withComment,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageMetadataKinds.Comments,
                result.Metadata.Lost & OfficeImageMetadataKinds.Comments);
        }

        [Fact]
        public void OfficeImageOptimizerSwapsDensityForExifOrientationBetweenProgressiveScans() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(2, 1, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    Progressive = true,
                    DpiX = 72D,
                    DpiY = 144D
                });
            byte[] exifPayload = System.Text.Encoding.ASCII.GetBytes("Exif\0\0")
                .Concat(CreateExifOrientation(6))
                .ToArray();
            byte[] oriented = InsertJpegSegmentBeforeSecondScan(jpeg, 0xE1, exifPayload);

            Assert.True(OfficeJpegCodec.TryDecode(oriented, out OfficeRasterImage? decoded));
            Assert.Equal((1, 2), (decoded!.Width, decoded.Height));
            Assert.True(OfficeImageOrientationNormalizer.TryRead(oriented, out OfficeImageOrientation orientation));
            Assert.Equal(OfficeImageOrientation.Rotate90Clockwise, orientation);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                oriented,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.InRange(result.Final.DpiX, 143.98D, 144.02D);
            Assert.InRange(result.Final.DpiY, 71.98D, 72.02D);
        }

        [Fact]
        public void OfficeImageOptimizerBoundsInputBeforeIdentificationAndMetadataInspection() {
            OfficeImageOptimizer.ValidateInputLength(OfficeRasterGuards.MaximumEncodedBytes);
            Assert.Throws<ArgumentException>(() =>
                OfficeImageOptimizer.ValidateInputLength(OfficeRasterGuards.MaximumEncodedBytes + 1));
        }

        [Fact]
        public void OfficeImageOptimizerSwapsInheritedDensityAxesAfterExifRotation() {
            byte[] jpeg = OfficeJpegCodec.Encode(
                new OfficeRasterImage(2, 1, OfficeColor.SteelBlue),
                new OfficeJpegEncodeOptions {
                    DpiX = 72D,
                    DpiY = 144D,
                    Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
                });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Jpeg,
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.Equal((1, 1), (result.Final.Width, result.Final.Height));
            Assert.InRange(result.Final.DpiX, 143.98D, 144.02D);
            Assert.InRange(result.Final.DpiY, 71.98D, 72.02D);
        }

        [Fact]
        public void OfficeImageOptimizerDoesNotRestoreMetadataWhenARequiredStripRewriteIsLarger() {
            var source = new OfficeRasterImage(32, 32);
            for (int y = 0; y < source.Height; y++) {
                for (int x = 0; x < source.Width; x++) {
                    source.SetPixel(x, y, OfficeColor.FromRgba(
                        (byte)(x * 17 + y * 11),
                        (byte)(x * 7 + y * 19),
                        (byte)(x * 23 + y * 3),
                        255));
                }
            }
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 20,
                Subsampling = OfficeJpegSubsampling.Y420,
                Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(3))
            });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                jpeg,
                new OfficeImageOptimizationRequest(source.Width, source.Height) {
                    OutputFormat = OfficeImageFormat.Png,
                    MetadataPolicy = OfficeImageMetadataPolicy.Strip
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageFormat.Png, result.Final.Format);
            Assert.True(result.FinalEncodedLength > result.OriginalEncodedLength);
            Assert.True(result.Metadata.PolicyApplied);
            Assert.NotEqual(OfficeImageMetadataKinds.None, result.Metadata.Stripped);
            Assert.False(ContainsSequence(result.Bytes, CreateExifOrientation(3)));
        }

        [Fact]
        public void OfficeImageOrientationNormalizer_AppliesOrIgnoresExifWithoutPlatformCodecs() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Blue);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 100,
                Subsampling = OfficeJpegSubsampling.Y444,
                DpiX = 72D,
                DpiY = 144D,
                Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
            });

            Assert.True(OfficeImageOrientationNormalizer.TryRead(jpeg, out OfficeImageOrientation orientation));
            Assert.Equal(OfficeImageOrientation.Rotate90Clockwise, orientation);
            Assert.True(OfficeImageOrientationNormalizer.TryNormalizeToPng(jpeg, true, out byte[] orientedPng, out OfficeImageInfo? orientedInfo));
            Assert.True(OfficeImageOrientationNormalizer.TryNormalizeToPng(jpeg, false, out byte[] rawPng, out OfficeImageInfo? rawInfo));
            Assert.Equal((1, 2), (orientedInfo!.Width, orientedInfo.Height));
            Assert.Equal((2, 1), (rawInfo!.Width, rawInfo.Height));
            Assert.InRange(orientedInfo.DpiX, 143.9D, 144.1D);
            Assert.InRange(orientedInfo.DpiY, 71.9D, 72.1D);
            Assert.InRange(rawInfo.DpiX, 71.9D, 72.1D);
            Assert.InRange(rawInfo.DpiY, 143.9D, 144.1D);
            Assert.True(OfficePngReader.TryDecode(orientedPng, out OfficeRasterImage? oriented));
            Assert.True(OfficePngReader.TryDecode(rawPng, out OfficeRasterImage? raw));
            AssertColorNear(oriented!.GetPixel(0, 0), OfficeColor.Red, 12);
            AssertColorNear(oriented.GetPixel(0, 1), OfficeColor.Blue, 12);
            AssertColorNear(raw!.GetPixel(0, 0), OfficeColor.Red, 12);
            AssertColorNear(raw.GetPixel(1, 0), OfficeColor.Blue, 12);
        }

        [Fact]
        public void OfficeImageOrientationNormalizer_IgnoresExifSegmentWithoutOrientationBeforeOrientedSegment() {
            var source = new OfficeRasterImage(2, 1);
            source.SetPixel(0, 0, OfficeColor.Red);
            source.SetPixel(1, 0, OfficeColor.Blue);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 100,
                Subsampling = OfficeJpegSubsampling.Y444,
                Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
            });
            jpeg = InsertExifSegmentAfterStartOfImage(jpeg, CreateExifWithoutOrientation());

            Assert.True(OfficeImageOrientationNormalizer.TryRead(jpeg, out OfficeImageOrientation orientation));
            Assert.Equal(OfficeImageOrientation.Rotate90Clockwise, orientation);
            Assert.True(OfficeImageOrientationNormalizer.TryNormalizeToPng(jpeg, false, out byte[] rawPng, out OfficeImageInfo? rawInfo));
            Assert.Equal((2, 1), (rawInfo!.Width, rawInfo.Height));
            Assert.True(OfficePngReader.TryDecode(rawPng, out OfficeRasterImage? raw));
            AssertColorNear(raw!.GetPixel(0, 0), OfficeColor.Red, 12);
            AssertColorNear(raw.GetPixel(1, 0), OfficeColor.Blue, 12);
        }

        [Fact]
        public void OfficeJpegCodec_RejectsOrientedDecodeBeforeAllocatingSecondRgbaBuffer() {
            var source = new OfficeRasterImage(8, 8, OfficeColor.Red);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
            });
            SetJpegFrameDimensions(jpeg, 5000, 5000);

            FormatException exception = Assert.Throws<FormatException>(() => OfficeJpegCodec.Decode(jpeg));

            Assert.Contains("dimensions exceed", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OfficeJpegCodec_DecodesIndependentJpegFixture() {
            byte[] jpeg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Kulek.jpg"));
            OfficeImageInfo identified = OfficeImageReader.Identify(jpeg);

            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);
            Assert.Equal(OfficeImageFormat.Jpeg, identified.Format);
            Assert.Equal(identified.Width, decoded.Width);
            Assert.Equal(identified.Height, decoded.Height);
            Assert.True(decoded.Width > 100);
            Assert.True(decoded.Height > 100);
            using (SHA256 sha256 = SHA256.Create()) {
                Assert.Equal(
                    "EA87164A1FF1B2CE978E3C007382CD90AAAC5269078CBA79306FD35972231E0D",
                    BitConverter.ToString(sha256.ComputeHash(decoded.GetPixels())).Replace("-", string.Empty));
            }
        }

        [Fact]
        public void OfficeJpegCodec_DecodesBaselineRestartIntervalsWithoutFastPeekCrossingMarkers() {
            byte[] jpeg = Convert.FromBase64String(
                "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/" +
                "2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAAQACADAREAAhEBAxEB/" +
                "8QAFAABAAAAAAAAAAAAAAAAAAAACP/EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAVAQEBAAAAAAAAAAAAAAAAAAAHCf/EABQRAQAAAAAAAAAAAAAAAAAAAAD/" +
                "3QAEAAH/2gAMAwEAAhEDEQA/ADoDFU3/0DoDFU3/0ToDFU3/0joDFU3/0zoDFU3/1DoDFU3/1ToDFU3/1joDFU3/2Q==");

            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);

            Assert.Equal(32, decoded.Width);
            Assert.Equal(16, decoded.Height);
            AssertColorNear(decoded.GetPixel(0, 0), OfficeColor.Red, 8);
            AssertColorNear(decoded.GetPixel(31, 15), OfficeColor.Red, 8);
        }

        [Fact]
        public void OfficeJpegCodec_FlattensTransparencyAgainstConfiguredBackground() {
            var source = new OfficeRasterImage(8, 8, OfficeColor.FromRgba(255, 0, 0, 128));

            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
                Quality = 100,
                Background = OfficeColor.Blue
            });
            OfficeRasterImage decoded = OfficeJpegCodec.Decode(jpeg);

            AssertColorNear(decoded.GetPixel(4, 4), OfficeColor.FromRgb(128, 0, 127), 12);
        }

        [Fact]
        public void OfficeImageOptimizer_DownsamplesUsingPlacementBounds() {
            OfficeRasterImage source = CreateQuadrantImage(160, 120);
            byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions { Quality = 94 });
            var request = new OfficeImageOptimizationRequest(40, 40) {
                JpegQuality = 80,
                KeepOriginalWhenNotSmaller = false
            };

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(jpeg, request);

            Assert.True(result.Changed);
            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(160, result.Original.Width);
            Assert.Equal(120, result.Original.Height);
            Assert.Equal(40, result.Final.Width);
            Assert.Equal(30, result.Final.Height);
            Assert.Equal(OfficeImageFormat.Jpeg, result.Final.Format);
            Assert.Equal(jpeg.LongLength, result.OriginalEncodedLength);
            Assert.Equal(result.Bytes.LongLength, result.FinalEncodedLength);
            Assert.True(OfficeJpegCodec.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.Equal(40, decoded!.Width);
            Assert.Equal(30, decoded.Height);
        }

        [Fact]
        public void OfficeImageOptimizer_DownsamplesPngWithoutLosingAlpha() {
            var source = new OfficeRasterImage(80, 40, OfficeColor.FromRgba(20, 80, 200, 96));
            byte[] png = OfficePngWriter.Encode(source);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(20, 20) {
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.True(result.Changed);
            Assert.Equal(OfficeImageFormat.Png, result.Final.Format);
            Assert.Equal(20, result.Final.Width);
            Assert.Equal(10, result.Final.Height);
            Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.InRange(decoded!.GetPixel(10, 5).A, 95, 97);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Png)]
        [InlineData(OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageFormat.Webp)]
        public void OfficeImageOptimizer_EncodesRequestedFormatWithSourceResolution(OfficeImageFormat outputFormat) {
            var source = new OfficeRasterImage(16, 12, OfficeColor.SteelBlue);
            byte[] png = OfficePngWriter.Encode(source, new OfficePngEncodeOptions {
                DpiX = 144D,
                DpiY = 120D
            });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(8, 6) {
                    OutputFormat = outputFormat,
                    KeepOriginalWhenNotSmaller = false,
                    TiffCompression = OfficeTiffCompression.Deflate,
                    JpegOptimizeHuffman = true
                });

            OfficeImageInfo encodedInfo = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(outputFormat, result.Final.Format);
            Assert.Equal(outputFormat, encodedInfo.Format);
            Assert.Equal((8, 6), (encodedInfo.Width, encodedInfo.Height));
            Assert.InRange(encodedInfo.DpiX, 143.98D, 144.02D);
            Assert.InRange(encodedInfo.DpiY, 119.98D, 120.02D);
            Assert.Equal(encodedInfo.DpiX, result.Final.DpiX);
            Assert.Equal(encodedInfo.DpiY, result.Final.DpiY);
        }

        [Fact]
        public void OfficeImageOptimizerPreservesBmpPhysicalResolutionAndReportsIt() {
            byte[] bmp = CreateBmp24WithResolution(5669, 4724);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                bmp,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Source & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Lost & OfficeImageMetadataKinds.Resolution);
            Assert.InRange(result.Final.DpiX, 143.98D, 144.02D);
            Assert.InRange(result.Final.DpiY, 119.98D, 120.02D);
        }

        [Fact]
        public void OfficeImageOptimizerReportsUnitlessPngResolutionAsLost() {
            byte[] png = OfficePngWriter.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficePngEncodeOptions { DpiX = 144D, DpiY = 72D });
            int physicalResolution = FindPngChunk(png, "pHYs");
            png[physicalResolution + 16] = 0;
            WritePngChunkCrc(png, physicalResolution, 9);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Png,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Source & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Lost & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void OfficeImageOptimizerReportsUnitlessTiffResolutionAsLost() {
            byte[] tiff = OfficeTiffCodec.Encode(
                new OfficeRasterImage(4, 4, OfficeColor.SteelBlue),
                new OfficeTiffEncodeOptions { DpiX = 300D, DpiY = 150D });
            int resolutionUnit = FindClassicTiffEntry(tiff, 296);
            WriteLittleEndian(tiff, resolutionUnit + 8, 1);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                tiff,
                new OfficeImageOptimizationRequest(2, 2) {
                    OutputFormat = OfficeImageFormat.Tiff,
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Preserve
                });

            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Source & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.None,
                result.Metadata.Preserved & OfficeImageMetadataKinds.Resolution);
            Assert.Equal(OfficeImageMetadataKinds.Resolution,
                result.Metadata.Lost & OfficeImageMetadataKinds.Resolution);
        }

        [Fact]
        public void ExplicitOutputResolutionForcesRewriteWhenOriginalIsSmaller() {
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.SteelBlue));

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(1, 1) {
                    OutputFormat = OfficeImageFormat.Png,
                    OutputDpiX = 300D,
                    OutputDpiY = 150D,
                    KeepOriginalWhenNotSmaller = true
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.NotEqual(png, result.Bytes);
            Assert.InRange(result.Final.DpiX, 299.98D, 300.02D);
            Assert.InRange(result.Final.DpiY, 149.98D, 150.02D);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Jpeg, 65535D)]
        [InlineData(OfficeImageFormat.Tiff, 1000000D)]
        [InlineData(OfficeImageFormat.Webp, 1000000D)]
        public void OfficeImageOptimizer_ClampsSourceResolutionToDestinationMaximum(
            OfficeImageFormat outputFormat,
            double expectedMaximumDpi) {
            var source = new OfficeRasterImage(4, 4, OfficeColor.SteelBlue);
            byte[] png = OfficePngWriter.Encode(source, new OfficePngEncodeOptions {
                DpiX = 60000000D,
                DpiY = 60000000D
            });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                png,
                new OfficeImageOptimizationRequest(4, 4) {
                    OutputFormat = outputFormat,
                    KeepOriginalWhenNotSmaller = false
                });

            OfficeImageInfo encodedInfo = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.InRange(encodedInfo.DpiX, expectedMaximumDpi - 0.01D, expectedMaximumDpi + 0.01D);
            Assert.InRange(encodedInfo.DpiY, expectedMaximumDpi - 0.01D, expectedMaximumDpi + 0.01D);
            Assert.Equal(encodedInfo.DpiX, result.Final.DpiX);
            Assert.Equal(encodedInfo.DpiY, result.Final.DpiY);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Png, 0.0127D)]
        [InlineData(OfficeImageFormat.Jpeg, 0.5D)]
        public void OfficeImageOptimizer_ClampsSourceResolutionToDestinationMinimum(
            OfficeImageFormat outputFormat,
            double expectedMinimumDpi) {
            var source = new OfficeRasterImage(4, 4, OfficeColor.SteelBlue);
            byte[] webp = OfficeWebpCodec.Encode(source, 0.0001D, 0.0001D);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                webp,
                new OfficeImageOptimizationRequest(4, 4) {
                    OutputFormat = outputFormat,
                    KeepOriginalWhenNotSmaller = false
                });

            OfficeImageInfo encodedInfo = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.True(encodedInfo.DpiX >= expectedMinimumDpi);
            Assert.True(encodedInfo.DpiY >= expectedMinimumDpi);
            Assert.Equal(encodedInfo.DpiX, result.Final.DpiX);
            Assert.Equal(encodedInfo.DpiY, result.Final.DpiY);
        }

        [Fact]
        public void OfficeImageOptimizer_UsesExplicitOutputResolutionWhenNoResizeIsNeeded() {
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.SteelBlue));
            var request = new OfficeImageOptimizationRequest(4, 4) {
                OutputDpiX = 300D,
                OutputDpiY = 150D,
                KeepOriginalWhenNotSmaller = false
            };

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(png, request);

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.InRange(result.Final.DpiX, 299.98D, 300.02D);
            Assert.InRange(result.Final.DpiY, 149.98D, 150.02D);
        }

        [Fact]
        public void OfficeImageOptimizationRequest_RejectsInvalidOutputResolution() {
            var request = new OfficeImageOptimizationRequest(4, 4);

            Assert.Throws<ArgumentOutOfRangeException>(() => request.OutputDpiX = 0D);
            Assert.Throws<ArgumentOutOfRangeException>(() => request.OutputDpiX = double.NaN);
            Assert.Throws<ArgumentOutOfRangeException>(() => request.OutputDpiY = double.PositiveInfinity);
        }

        [Theory]
        [InlineData(OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageFormat.Webp)]
        public void OfficeImageOptimizer_ConvertsManagedStaticInputsToPng(OfficeImageFormat sourceFormat) {
            OfficeRasterImage source = CreateQuadrantImage(16, 12);
            OfficeImageExportFormat exportFormat = sourceFormat == OfficeImageFormat.Tiff
                ? OfficeImageExportFormat.Tiff
                : OfficeImageExportFormat.Webp;
            byte[] encoded = OfficeRasterImageEncoder.Encode(source, exportFormat);

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                encoded,
                new OfficeImageOptimizationRequest(8, 6) {
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(sourceFormat, result.Original.Format);
            Assert.Equal(OfficeImageFormat.Png, result.Final.Format);
            Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.Equal((8, 6), (decoded!.Width, decoded.Height));
        }

        [Fact]
        public void OfficeImageOptimizer_ConvertsStaticGifWithoutImplicitAnimationLoss() {
            byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                gif,
                new OfficeImageOptimizationRequest(1, 1) {
                    KeepOriginalWhenNotSmaller = false
                });

            Assert.Equal(OfficeImageOptimizationStatus.Optimized, result.Status);
            Assert.Equal(OfficeImageFormat.Gif, result.Original.Format);
            Assert.Equal(OfficeImageFormat.Png, result.Final.Format);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? decoded));
            Assert.Equal((1, 1), (decoded!.Width, decoded.Height));
        }

        [Fact]
        public void OfficeImageOptimizerRejectsMultiPageTiffWithoutExplicitPageLoss() {
            byte[] tiff = OfficeTiffCodec.EncodePages(new[] {
                new OfficeRasterImage(4, 3, OfficeColor.Red),
                new OfficeRasterImage(2, 2, OfficeColor.Blue)
            }, new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.Lzw });

            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                tiff,
                new OfficeImageOptimizationRequest(2, 2) {
                    KeepOriginalWhenNotSmaller = false,
                    MetadataPolicy = OfficeImageMetadataPolicy.Strip
                });

            Assert.Equal(OfficeImageOptimizationStatus.DecodeFailed, result.Status);
            Assert.False(result.Changed);
            Assert.Equal(tiff, result.Bytes);
            Assert.False(result.Metadata.PolicyApplied);
            Assert.Equal(OfficeImageMetadataKinds.None, result.Metadata.Stripped);
        }

        [Fact]
        public void OfficeImageOptimizer_ResultKeepsEncodedBytesImmutable() {
            byte[] source = OfficePngWriter.Encode(new OfficeRasterImage(8, 8, OfficeColor.SteelBlue));
            OfficeImageOptimizationResult result = OfficeImageOptimizer.Optimize(
                source,
                new OfficeImageOptimizationRequest(4, 4) {
                    KeepOriginalWhenNotSmaller = false
                });
            byte[] first = result.Bytes;
            byte expected = first[0];

            first[0] ^= 0xFF;

            Assert.Equal(expected, result.Bytes[0]);
        }

        private static OfficeRasterImage CreateQuadrantImage(int width, int height) {
            var image = new OfficeRasterImage(width, height);
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    OfficeColor color = x < width / 2
                        ? (y < height / 2 ? OfficeColor.Red : OfficeColor.Blue)
                        : (y < height / 2 ? OfficeColor.Lime : OfficeColor.White);
                    image.SetPixel(x, y, color);
                }
            }
            return image;
        }

        private static byte[] CreateBmp24WithResolution(int horizontalPixelsPerMeter, int verticalPixelsPerMeter) {
            const int width = 2;
            const int height = 1;
            const int pixelOffset = 54;
            const int rowStride = 8;
            var bmp = new byte[pixelOffset + rowStride];
            bmp[0] = (byte)'B';
            bmp[1] = (byte)'M';
            WriteLittleEndian(bmp, 2, bmp.Length);
            WriteLittleEndian(bmp, 10, pixelOffset);
            WriteLittleEndian(bmp, 14, 40);
            WriteLittleEndian(bmp, 18, width);
            WriteLittleEndian(bmp, 22, height);
            bmp[26] = 1;
            bmp[28] = 24;
            WriteLittleEndian(bmp, 34, rowStride);
            WriteLittleEndian(bmp, 38, horizontalPixelsPerMeter);
            WriteLittleEndian(bmp, 42, verticalPixelsPerMeter);
            bmp[pixelOffset + 2] = 255;
            bmp[pixelOffset + 3] = 255;
            return bmp;
        }

        private static int FindClassicTiffEntry(byte[] bytes, int expectedTag) {
            int ifdOffset = ReadLittleEndian(bytes, 4);
            int entryCount = bytes[ifdOffset] | bytes[ifdOffset + 1] << 8;
            for (int index = 0; index < entryCount; index++) {
                int entryOffset = ifdOffset + 2 + index * 12;
                int tag = bytes[entryOffset] | bytes[entryOffset + 1] << 8;
                if (tag == expectedTag) return entryOffset;
            }
            throw new InvalidOperationException("TIFF entry was not found.");
        }

        private static int FindPngChunk(byte[] bytes, string expectedType) {
            int offset = 8;
            while (offset <= bytes.Length - 12) {
                int length = bytes[offset] << 24 |
                             bytes[offset + 1] << 16 |
                             bytes[offset + 2] << 8 |
                             bytes[offset + 3];
                if (System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4) == expectedType) return offset;
                offset += length + 12;
            }
            throw new InvalidOperationException("PNG chunk was not found.");
        }

        private static void WritePngChunkCrc(byte[] bytes, int chunkOffset, int length) {
            uint crc = 0xFFFFFFFFU;
            for (int index = chunkOffset + 4; index < chunkOffset + 8 + length; index++) {
                crc ^= bytes[index];
                for (int bit = 0; bit < 8; bit++) {
                    crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
                }
            }
            crc ^= 0xFFFFFFFFU;
            int offset = chunkOffset + 8 + length;
            bytes[offset] = (byte)(crc >> 24);
            bytes[offset + 1] = (byte)(crc >> 16);
            bytes[offset + 2] = (byte)(crc >> 8);
            bytes[offset + 3] = (byte)crc;
        }

        private static int ReadLittleEndian(byte[] bytes, int offset) =>
            bytes[offset] |
            bytes[offset + 1] << 8 |
            bytes[offset + 2] << 16 |
            bytes[offset + 3] << 24;

        private static void WriteLittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static byte[] CreateExifOrientation(ushort orientation) => new byte[] {
            (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            0x01, 0x00,
            0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
            (byte)orientation, (byte)(orientation >> 8), 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        };

        private static byte[] CreateMinimalOptimizationIccProfile(string deviceColorSpace = "RGB ") {
            Assert.Equal(4, deviceColorSpace.Length);
            var profile = new byte[132];
            profile[0] = 0;
            profile[1] = 0;
            profile[2] = 0;
            profile[3] = 132;
            profile[16] = (byte)deviceColorSpace[0];
            profile[17] = (byte)deviceColorSpace[1];
            profile[18] = (byte)deviceColorSpace[2];
            profile[19] = (byte)deviceColorSpace[3];
            profile[36] = (byte)'a';
            profile[37] = (byte)'c';
            profile[38] = (byte)'s';
            profile[39] = (byte)'p';
            return profile;
        }

        private static byte[] CreateExifWithOrientationAndResolution(
            ushort orientation,
            int dpiX,
            int dpiY) {
            var exif = new byte[78];
            exif[0] = (byte)'I';
            exif[1] = (byte)'I';
            exif[2] = 0x2A;
            exif[4] = 0x08;
            exif[8] = 0x04;
            WriteLittleEndianUInt16(exif, 10, 274);
            WriteLittleEndianUInt16(exif, 12, 3);
            WriteLittleEndianUInt32(exif, 14, 1);
            WriteLittleEndianUInt16(exif, 18, orientation);
            WriteLittleEndianUInt16(exif, 22, 282);
            WriteLittleEndianUInt16(exif, 24, 5);
            WriteLittleEndianUInt32(exif, 26, 1);
            WriteLittleEndianUInt32(exif, 30, 62);
            WriteLittleEndianUInt16(exif, 34, 283);
            WriteLittleEndianUInt16(exif, 36, 5);
            WriteLittleEndianUInt32(exif, 38, 1);
            WriteLittleEndianUInt32(exif, 42, 70);
            WriteLittleEndianUInt16(exif, 46, 296);
            WriteLittleEndianUInt16(exif, 48, 3);
            WriteLittleEndianUInt32(exif, 50, 1);
            WriteLittleEndianUInt16(exif, 54, 2);
            WriteLittleEndianUInt32(exif, 62, checked((uint)dpiX));
            WriteLittleEndianUInt32(exif, 66, 1);
            WriteLittleEndianUInt32(exif, 70, checked((uint)dpiY));
            WriteLittleEndianUInt32(exif, 74, 1);
            return exif;
        }

        private static bool ContainsSequence(byte[] container, byte[] value) {
            if (value.Length == 0 || value.Length > container.Length) return false;
            for (int offset = 0; offset <= container.Length - value.Length; offset++) {
                int index = 0;
                while (index < value.Length && container[offset + index] == value[index]) index++;
                if (index == value.Length) return true;
            }
            return false;
        }

        private static byte[] CreateExifWithoutOrientation() => new byte[] {
            (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        };

        private static byte[] CreateExifWithResolution(int dpiX = 72, int dpiY = 72) {
            byte[] exif = new byte[] {
            (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            0x03, 0x00,
            0x1A, 0x01, 0x05, 0x00, 0x01, 0x00, 0x00, 0x00, 0x32, 0x00, 0x00, 0x00,
            0x1B, 0x01, 0x05, 0x00, 0x01, 0x00, 0x00, 0x00, 0x3A, 0x00, 0x00, 0x00,
            0x28, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00
            };
            WriteLittleEndianUInt32(exif, 50, checked((uint)dpiX));
            WriteLittleEndianUInt32(exif, 58, checked((uint)dpiY));
            return exif;
        }

        private static byte[] CreateExifWithResolutionAliasedToWhitePoint() {
            var exif = new byte[88];
            exif[0] = (byte)'I';
            exif[1] = (byte)'I';
            WriteLittleEndianUInt16(exif, 2, 42);
            WriteLittleEndianUInt32(exif, 4, 8);
            WriteLittleEndianUInt16(exif, 8, 4);

            WriteLittleEndianUInt16(exif, 10, 282);
            WriteLittleEndianUInt16(exif, 12, 5);
            WriteLittleEndianUInt32(exif, 14, 1);
            WriteLittleEndianUInt32(exif, 18, 64);

            WriteLittleEndianUInt16(exif, 22, 283);
            WriteLittleEndianUInt16(exif, 24, 5);
            WriteLittleEndianUInt32(exif, 26, 1);
            WriteLittleEndianUInt32(exif, 30, 80);

            WriteLittleEndianUInt16(exif, 34, 296);
            WriteLittleEndianUInt16(exif, 36, 3);
            WriteLittleEndianUInt32(exif, 38, 1);
            WriteLittleEndianUInt16(exif, 42, 2);

            WriteLittleEndianUInt16(exif, 46, 318);
            WriteLittleEndianUInt16(exif, 48, 5);
            WriteLittleEndianUInt32(exif, 50, 2);
            WriteLittleEndianUInt32(exif, 54, 64);

            WriteLittleEndianUInt32(exif, 64, 72);
            WriteLittleEndianUInt32(exif, 68, 1);
            WriteLittleEndianUInt32(exif, 72, 1);
            WriteLittleEndianUInt32(exif, 76, 1);
            WriteLittleEndianUInt32(exif, 80, 96);
            WriteLittleEndianUInt32(exif, 84, 1);
            return exif;
        }

        private static byte[] CreateExifWithResolutionAliasedToLongSubIfd() {
            var exif = new byte[88];
            exif[0] = (byte)'I';
            exif[1] = (byte)'I';
            WriteLittleEndianUInt16(exif, 2, 42);
            WriteLittleEndianUInt32(exif, 4, 8);
            WriteLittleEndianUInt16(exif, 8, 4);

            WriteLittleEndianUInt16(exif, 10, 282);
            WriteLittleEndianUInt16(exif, 12, 5);
            WriteLittleEndianUInt32(exif, 14, 1);
            WriteLittleEndianUInt32(exif, 18, 64);

            WriteLittleEndianUInt16(exif, 22, 283);
            WriteLittleEndianUInt16(exif, 24, 5);
            WriteLittleEndianUInt32(exif, 26, 1);
            WriteLittleEndianUInt32(exif, 30, 80);

            WriteLittleEndianUInt16(exif, 34, 296);
            WriteLittleEndianUInt16(exif, 36, 3);
            WriteLittleEndianUInt32(exif, 38, 1);
            WriteLittleEndianUInt16(exif, 42, 2);

            WriteLittleEndianUInt16(exif, 46, 330);
            WriteLittleEndianUInt16(exif, 48, 4);
            WriteLittleEndianUInt32(exif, 50, 1);
            WriteLittleEndianUInt32(exif, 54, 64);

            WriteLittleEndianUInt32(exif, 80, 96);
            WriteLittleEndianUInt32(exif, 84, 1);
            return exif;
        }

        private static byte[] CreateExifWithOrientationAliasedToWhitePoint() {
            var exif = new byte[40];
            exif[0] = (byte)'I';
            exif[1] = (byte)'I';
            WriteLittleEndianUInt16(exif, 2, 42);
            WriteLittleEndianUInt32(exif, 4, 8);
            WriteLittleEndianUInt16(exif, 8, 2);

            WriteLittleEndianUInt16(exif, 10, 274);
            WriteLittleEndianUInt16(exif, 12, 3);
            WriteLittleEndianUInt32(exif, 14, 1);
            WriteLittleEndianUInt16(exif, 18, 6);

            WriteLittleEndianUInt16(exif, 22, 318);
            WriteLittleEndianUInt16(exif, 24, 5);
            WriteLittleEndianUInt32(exif, 26, 2);
            WriteLittleEndianUInt32(exif, 30, 18);
            return exif;
        }

        private static byte[] CreateExifWithDuplicateOrientations() {
            var exif = new byte[38];
            exif[0] = (byte)'I';
            exif[1] = (byte)'I';
            WriteLittleEndianUInt16(exif, 2, 42);
            WriteLittleEndianUInt32(exif, 4, 8);
            WriteLittleEndianUInt16(exif, 8, 2);

            WriteLittleEndianUInt16(exif, 10, 274);
            WriteLittleEndianUInt16(exif, 12, 3);
            WriteLittleEndianUInt32(exif, 14, 1);
            WriteLittleEndianUInt16(exif, 18, 6);

            WriteLittleEndianUInt16(exif, 22, 274);
            WriteLittleEndianUInt16(exif, 24, 3);
            WriteLittleEndianUInt32(exif, 26, 1);
            WriteLittleEndianUInt16(exif, 30, 3);
            return exif;
        }

        private static void WriteLittleEndianUInt32(byte[] data, int offset, uint value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteLittleEndianUInt16(byte[] data, int offset, ushort value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        }

        private static byte[] InsertExifSegmentAfterStartOfImage(byte[] jpeg, byte[] tiffData) {
            int payloadLength = 6 + tiffData.Length;
            int segmentLength = payloadLength + 2;
            var segment = new byte[segmentLength + 2];
            segment[0] = 0xFF;
            segment[1] = 0xE1;
            segment[2] = (byte)(segmentLength >> 8);
            segment[3] = (byte)segmentLength;
            segment[4] = (byte)'E';
            segment[5] = (byte)'x';
            segment[6] = (byte)'i';
            segment[7] = (byte)'f';
            Buffer.BlockCopy(tiffData, 0, segment, 10, tiffData.Length);

            var combined = new byte[jpeg.Length + segment.Length];
            Buffer.BlockCopy(jpeg, 0, combined, 0, 2);
            Buffer.BlockCopy(segment, 0, combined, 2, segment.Length);
            Buffer.BlockCopy(jpeg, 2, combined, 2 + segment.Length, jpeg.Length - 2);
            return combined;
        }

        private static byte[] InsertApp1SegmentAfterStartOfImage(byte[] jpeg, byte[] payload) {
            int segmentLength = payload.Length + 2;
            var segment = new byte[segmentLength + 2];
            segment[0] = 0xFF;
            segment[1] = 0xE1;
            segment[2] = (byte)(segmentLength >> 8);
            segment[3] = (byte)segmentLength;
            Buffer.BlockCopy(payload, 0, segment, 4, payload.Length);
            var combined = new byte[jpeg.Length + segment.Length];
            Buffer.BlockCopy(jpeg, 0, combined, 0, 2);
            Buffer.BlockCopy(segment, 0, combined, 2, segment.Length);
            Buffer.BlockCopy(jpeg, 2, combined, 2 + segment.Length, jpeg.Length - 2);
            return combined;
        }

        private static byte[] InsertJpegSegmentBeforeSecondScan(byte[] jpeg, byte marker, byte[] payload) {
            int scanCount = 0;
            int insertionOffset = -1;
            for (int index = 0; index < jpeg.Length - 1; index++) {
                if (jpeg[index] != 0xFF || jpeg[index + 1] != 0xDA) continue;
                if (++scanCount == 2) {
                    insertionOffset = index;
                    break;
                }
            }
            Assert.True(insertionOffset > 0);
            int segmentLength = checked(payload.Length + 2);
            var combined = new byte[jpeg.Length + payload.Length + 4];
            Buffer.BlockCopy(jpeg, 0, combined, 0, insertionOffset);
            combined[insertionOffset] = 0xFF;
            combined[insertionOffset + 1] = marker;
            combined[insertionOffset + 2] = (byte)(segmentLength >> 8);
            combined[insertionOffset + 3] = (byte)segmentLength;
            Buffer.BlockCopy(payload, 0, combined, insertionOffset + 4, payload.Length);
            Buffer.BlockCopy(jpeg, insertionOffset, combined, insertionOffset + payload.Length + 4,
                jpeg.Length - insertionOffset);
            return combined;
        }

        private static void SetJpegFrameDimensions(byte[] jpeg, ushort width, ushort height) {
            int frame = FindMarker(jpeg, 0xC0);
            if (frame < 0) frame = FindMarker(jpeg, 0xC2);
            Assert.True(frame >= 0);
            jpeg[frame + 5] = (byte)(height >> 8);
            jpeg[frame + 6] = (byte)height;
            jpeg[frame + 7] = (byte)(width >> 8);
            jpeg[frame + 8] = (byte)width;
        }

        private static byte[] BuildSeparateComponentBaselineJpeg() {
            byte[] seed = OfficeJpegCodec.Encode(
                new OfficeRasterImage(1, 1, OfficeColor.Red),
                new OfficeJpegEncodeOptions { Subsampling = OfficeJpegSubsampling.Y444 });
            int firstScan = FindMarker(seed, 0xDA);
            Assert.True(firstScan > 0);

            using var stream = new MemoryStream();
            stream.Write(seed, 0, firstScan);
            WriteBaselineScan(stream, componentId: 1, tableSelectors: 0x00, entropyByte: 0x2B);
            WriteBaselineScan(stream, componentId: 2, tableSelectors: 0x11, entropyByte: 0x0F);
            WriteBaselineScan(stream, componentId: 3, tableSelectors: 0x11, entropyByte: 0x0F);
            stream.WriteByte(0xFF);
            stream.WriteByte(0xD9);
            return stream.ToArray();
        }

        private static void WriteBaselineScan(Stream stream, byte componentId, byte tableSelectors, byte entropyByte) {
            byte[] scan = {
                0xFF, 0xDA, 0x00, 0x08, 0x01,
                componentId, tableSelectors,
                0x00, 0x3F, 0x00,
                entropyByte
            };
            stream.Write(scan, 0, scan.Length);
        }

        private static List<(int ComponentCount, int SpectralStart)> ReadStartOfScanHeaders(byte[] jpeg) {
            var scans = new List<(int ComponentCount, int SpectralStart)>();
            for (int index = 0; index + 5 < jpeg.Length; index++) {
                if (jpeg[index] != 0xFF || jpeg[index + 1] != 0xDA) continue;
                int componentCount = jpeg[index + 4];
                int spectralStartIndex = index + 5 + componentCount * 2;
                Assert.True(spectralStartIndex < jpeg.Length);
                scans.Add((componentCount, jpeg[spectralStartIndex]));
            }
            return scans;
        }

        private static int FindMarker(byte[] jpeg, byte marker) {
            int offset = 2;
            while (offset + 3 < jpeg.Length) {
                if (jpeg[offset] != 0xFF) {
                    offset++;
                    continue;
                }

                while (offset < jpeg.Length && jpeg[offset] == 0xFF) offset++;
                if (offset >= jpeg.Length) break;
                byte current = jpeg[offset++];
                if (current == marker) return offset - 2;
                if (current == 0xD9 || (current >= 0xD0 && current <= 0xD7)) continue;
                if (offset + 1 >= jpeg.Length) break;
                int length = (jpeg[offset] << 8) | jpeg[offset + 1];
                offset += length;
            }
            return -1;
        }

        private static void AssertColorNear(OfficeColor actual, OfficeColor expected, int tolerance) {
            Assert.InRange((int)actual.R, Math.Max(0, expected.R - tolerance), Math.Min(255, expected.R + tolerance));
            Assert.InRange((int)actual.G, Math.Max(0, expected.G - tolerance), Math.Min(255, expected.G + tolerance));
            Assert.InRange((int)actual.B, Math.Max(0, expected.B - tolerance), Math.Min(255, expected.B + tolerance));
            Assert.Equal(255, actual.A);
        }
    }
}
