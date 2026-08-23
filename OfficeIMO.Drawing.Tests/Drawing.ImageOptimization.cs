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

        private static byte[] CreateExifOrientation(ushort orientation) => new byte[] {
            (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            0x01, 0x00,
            0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
            (byte)orientation, (byte)(orientation >> 8), 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        };

        private static byte[] CreateExifWithoutOrientation() => new byte[] {
            (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        };

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
