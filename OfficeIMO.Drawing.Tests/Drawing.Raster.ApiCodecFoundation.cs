using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class RasterCodecFoundationTests {
        [Theory]
        [InlineData("horizontal", 0.000000000001)]
        [InlineData("horizontal", 4294967296D)]
        [InlineData("vertical", 0.000000000001)]
        [InlineData("vertical", 4294967296D)]
        [InlineData("typed", 0.000000000001)]
        [InlineData("typed", 4294967296D)]
        public void RejectedDensityAssignmentPreservesResolutionFieldsAndRemovalIntent(string route, double invalid) {
            byte[] encoded = OfficeTiffCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue));
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
            metadata.SetExifValue(OfficeExifTag.Artist, "Pending edit");
            metadata.RemoveExifValue(OfficeExifTag.YResolution);
            OfficeImageResolution beforeResolution = metadata.Resolution;
            byte[] beforeProfile = metadata.ExifProfile!;
            Assert.Throws<ArgumentOutOfRangeException>(() => {
                if (route == "horizontal") metadata.HorizontalResolution = invalid;
                else if (route == "vertical") metadata.VerticalResolution = invalid;
                else metadata.Resolution = new OfficeImageResolution(144D, invalid);
            });
            Assert.Same(beforeResolution, metadata.Resolution);
            Assert.Equal(beforeProfile, metadata.ExifProfile);
            Assert.Null(metadata.GetExifValue(OfficeExifTag.YResolution));
            OfficeImageMetadata rewritten = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, metadata.Clone()));
            Assert.Null(rewritten.GetExifValue(OfficeExifTag.YResolution));
            Assert.Equal("Pending edit", rewritten.GetExifValue(OfficeExifTag.Artist)!.Value);
        }

        [Fact]
        public void RejectedUnitChangePreservesNativeDensityAndExactRationals() {
            double density = 1D / uint.MaxValue;
            byte[] encoded = OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue),
                new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(density, density) });
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
            OfficeImageResolution beforeResolution = metadata.Resolution;
            byte[] beforeProfile = metadata.ExifProfile!;
            Assert.Throws<ArgumentOutOfRangeException>(() => metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter);
            Assert.Same(beforeResolution, metadata.Resolution);
            Assert.Equal(beforeProfile, metadata.ExifProfile);
            Assert.Equal(density, OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, metadata)).HorizontalResolution);
        }

        [Theory]
        [InlineData("x")]
        [InlineData("y")]
        [InlineData("unit")]
        [InlineData("profile")]
        public void RejectedGenericDensityOrProfileImportPreservesPendingMetadata(string route) {
            byte[] encoded = OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue));
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
            metadata.SetExifValue(OfficeExifTag.Artist, "Pending edit");
            metadata.RemoveExifValue(OfficeExifTag.YResolution);
            OfficeImageResolution beforeResolution = metadata.Resolution;
            byte[] beforeProfile = metadata.ExifProfile!;
            if (route == "profile") {
                Assert.Throws<FormatException>(() => metadata.ExifProfile = new byte[] { 1, 2, 3 });
            } else {
                Assert.ThrowsAny<ArgumentException>(() => {
                    if (route == "unit") metadata.SetExifValue(OfficeExifTag.ResolutionUnit, (ushort)4);
                    else metadata.SetExifValue(route == "x" ? OfficeExifTag.XResolution : OfficeExifTag.YResolution, new OfficeRational(0, 1));
                });
            }
            Assert.Same(beforeResolution, metadata.Resolution);
            Assert.Equal(beforeProfile, metadata.ExifProfile);
            Assert.Null(metadata.GetExifValue(OfficeExifTag.YResolution));
            Assert.Equal("Pending edit", metadata.GetExifValue(OfficeExifTag.Artist)!.Value);
        }
        [Fact]
        public void GenericExifDensityFromMeterCarrierUsesExifUnitsAndTypedAssignmentSynchronizesPartialProfiles() {
            byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue), OfficeImageExportFormat.Png);
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
            Assert.Equal(OfficeImageResolutionUnit.PixelsPerMeter, metadata.ResolutionUnits);
            double oldVerticalDpi = metadata.PhysicalDpiY!.Value;
            metadata.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(300, 1));
            Assert.Equal(OfficeImageResolutionUnit.PixelsPerInch, metadata.ResolutionUnits);
            Assert.Equal(300D, metadata.PhysicalDpiX);
            Assert.Equal(oldVerticalDpi, metadata.PhysicalDpiY);
            OfficeImageMetadata parsed = OfficeImageMetadata.ParseExifProfile(metadata.ExifProfile!);
            Assert.Equal(300D, parsed.Resolution.Horizontal);
            metadata.Resolution = new OfficeImageResolution(60D, 50D, OfficeImageResolutionUnit.PixelsPerCentimeter);
            parsed = OfficeImageMetadata.ParseExifProfile(metadata.ExifProfile!);
            Assert.Equal(OfficeImageResolutionUnit.PixelsPerCentimeter, parsed.ResolutionUnits);
            Assert.Equal(60D, parsed.Resolution.Horizontal); Assert.Equal(50D, parsed.Resolution.Vertical);
            OfficeImageMetadata applied = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, metadata));
            Assert.InRange(applied.PhysicalDpiX!.Value, 152.39D, 152.41D);
        }
        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        public void ImportedDensityTypedEditsCloneAndExplicitRemovalUseOneAuthority(OfficeImageExportFormat format) {
            byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), format);
            var imported = new OfficeImageMetadata();
            imported.SetExifValue(OfficeExifTag.Artist, "Density import");
            imported.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(300, 1));
            imported.SetExifValue(OfficeExifTag.YResolution, new OfficeRational(150, 1));
            imported.SetExifValue(OfficeExifTag.ResolutionUnit, (ushort)2);
            OfficeImageMetadata parsed = OfficeImageMetadata.ParseExifProfile(imported.EncodeExifProfile()!);
            Assert.Equal(300D, parsed.Resolution.Horizontal);
            var source = OfficeImageMetadata.Read(encoded);
            source.ExifProfile = imported.ExifProfile;
            Assert.Equal(150D, source.Resolution.Vertical);
            source.Resolution = new OfficeImageResolution(60D, 50D, OfficeImageResolutionUnit.PixelsPerCentimeter);
            Assert.Equal(60D, Assert.IsType<OfficeRational>(source.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
            Assert.Equal((ushort)3, source.GetExifValue(OfficeExifTag.ResolutionUnit)!.Value);
            OfficeImageMetadata cloned = source.Clone();
            cloned.RemoveExifValue(OfficeExifTag.XResolution);
            cloned.RemoveExifValue(OfficeExifTag.YResolution);
            cloned.RemoveExifValue(OfficeExifTag.ResolutionUnit);
            OfficeImageMetadata rewritten = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, cloned.Clone()));
            Assert.Null(rewritten.GetExifValue(OfficeExifTag.XResolution));
            Assert.Null(rewritten.GetExifValue(OfficeExifTag.YResolution));
            Assert.Null(rewritten.GetExifValue(OfficeExifTag.ResolutionUnit));
            Assert.NotNull(source.GetExifValue(OfficeExifTag.XResolution));
            cloned.Resolution = new OfficeImageResolution(72D, 144D);
            rewritten = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, cloned));
            Assert.Equal(72D, rewritten.PhysicalDpiX);
            Assert.Equal(144D, rewritten.PhysicalDpiY);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp)]
        [InlineData(OfficeImageExportFormat.Bmp)]
        public void MetadataAwareEncodingReportsActualOmissionsAndPreservesCallerState(OfficeImageExportFormat format) {
            var image = new OfficeRasterImage(3, 2, OfficeColor.Blue);
            var metadata = new OfficeImageMetadata { Resolution = new OfficeImageResolution(300D, 150D), IptcProfile = new byte[] { 0x1C, 2, 120, 0, 1, 65 } };
            metadata.SetExifValue(OfficeExifTag.Artist, "Author");
            var options = new OfficeRasterEncodingOptions { Resolution = new OfficeImageResolution(144D, 120D) };
            OfficeRasterEncodingResult result = OfficeRasterImageEncoder.EncodeWithMetadata(image, format, metadata, options);
            OfficeImageMetadata actual = OfficeImageMetadata.Read(result.EncodedBytes);
            Assert.InRange(actual.PhysicalDpiX!.Value, 143.98D, 144.02D);
            Assert.InRange(actual.PhysicalDpiY!.Value, 119.98D, 120.02D);
            Assert.Equal(300D, metadata.Resolution.Horizontal);
            Assert.Equal(144D, options.Resolution!.Horizontal);
            bool keepsIptc = format == OfficeImageExportFormat.Jpeg || format == OfficeImageExportFormat.Tiff;
            Assert.Equal(keepsIptc ? metadata.IptcProfile : null, actual.IptcProfile);
            Assert.Equal(!keepsIptc, result.HasMetadataLoss);
            if (keepsIptc) Assert.Same(result.EncodedBytes, result.RequireMetadataPreservation());
            else Assert.Throws<InvalidOperationException>(() => result.RequireMetadataPreservation());
            Assert.Equal(OfficeImageReader.Identify(result.EncodedBytes).Format, format.GetContainerFormat());
            Assert.Same(result.EncodedBytes, result.EncodedBytes);
            result.Metadata.Resolution = new OfficeImageResolution(72D, 72D);
            Assert.Equal(300D, metadata.Resolution.Horizontal);
            Assert.InRange(OfficeImageMetadata.Read(result.EncodedBytes).PhysicalDpiX!.Value, 143.98D, 144.02D);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp)]
        [InlineData(OfficeImageExportFormat.Bmp)]
        public void MetadataDensitySuppressionRemovesNativeAndExifCarriers(OfficeImageExportFormat format) {
            var metadata = new OfficeImageMetadata { Resolution = new OfficeImageResolution(300D, 150D) };
            metadata.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(300, 1));
            metadata.SetExifValue(OfficeExifTag.YResolution, new OfficeRational(150, 1));
            metadata.SetExifValue(OfficeExifTag.Artist, "Density suppression");
            var options = new OfficeRasterEncodingOptions { WriteResolutionMetadata = false };
            byte[] ordinary = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), format);
            OfficeRasterEncodingResult applied = metadata.ApplyForEncoding(ordinary, options);
            OfficeImageMetadata actual = OfficeImageMetadata.Read(applied.EncodedBytes);
            Assert.Null(actual.GetExifValue(OfficeExifTag.XResolution));
            Assert.Null(actual.GetExifValue(OfficeExifTag.YResolution));
            if (format == OfficeImageExportFormat.Png) Assert.DoesNotContain("pHYs", Encoding.ASCII.GetString(applied.EncodedBytes));
            if (format == OfficeImageExportFormat.Jpeg) Assert.DoesNotContain("JFIF", Encoding.ASCII.GetString(applied.EncodedBytes));
            if (format == OfficeImageExportFormat.Bmp) { Assert.Equal(0, BitConverter.ToInt32(applied.EncodedBytes, 38)); Assert.Equal(0, BitConverter.ToInt32(applied.EncodedBytes, 42)); }
            Assert.NotNull(metadata.GetExifValue(OfficeExifTag.XResolution));
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp)]
        public void NativeDensitySwitchRetainsDistinctExifCarriersWhereAvailable(OfficeImageExportFormat format) {
            var metadata = new OfficeImageMetadata();
            metadata.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(300, 1));
            metadata.SetExifValue(OfficeExifTag.YResolution, new OfficeRational(150, 1));
            metadata.SetExifValue(OfficeExifTag.ResolutionUnit, (ushort)2);
            metadata.SetExifValue(OfficeExifTag.Artist, "Carrier policy");
            var options = new OfficeRasterEncodingOptions();
            options.Png.WritePhysicalResolution = false; options.Jpeg.WriteJfifHeader = false;
            options.Tiff.WriteResolution = false; options.Webp.WritePhysicalResolution = false;
            OfficeRasterEncodingResult result = OfficeRasterImageEncoder.EncodeWithMetadata(new OfficeRasterImage(3, 2, OfficeColor.Blue), format, metadata, options);
            OfficeImageMetadata actual = OfficeImageMetadata.Read(result.EncodedBytes);
            if (format == OfficeImageExportFormat.Png || format == OfficeImageExportFormat.Jpeg) {
                Assert.Equal(300D, Assert.IsType<OfficeRational>(actual.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
                Assert.Equal(150D, Assert.IsType<OfficeRational>(actual.GetExifValue(OfficeExifTag.YResolution)!.Value).ToDouble());
            } else {
                Assert.Null(actual.GetExifValue(OfficeExifTag.XResolution));
                Assert.Null(actual.GetExifValue(OfficeExifTag.YResolution));
            }
            if (format == OfficeImageExportFormat.Png) Assert.DoesNotContain("pHYs", Encoding.ASCII.GetString(result.EncodedBytes));
            if (format == OfficeImageExportFormat.Jpeg) Assert.DoesNotContain("JFIF", Encoding.ASCII.GetString(result.EncodedBytes));
            Assert.Equal("Carrier policy", actual.GetExifValue(OfficeExifTag.Artist)!.Value);
        }

        [Fact]
        public void StillSequenceEncodingRetainsAllTiffPagesAndSourceEvidence() {
            var frames = new OfficeRasterFrames(new[] {
                new OfficeRasterFrame(new OfficeRasterImage(3, 2, OfficeColor.Red), TimeSpan.Zero),
                new OfficeRasterFrame(new OfficeRasterImage(2, 3, OfficeColor.Blue), TimeSpan.Zero)
            });
            var metadata = new OfficeImageMetadata { Resolution = new OfficeImageResolution(60D, 50D, OfficeImageResolutionUnit.PixelsPerCentimeter) };
            metadata.SetExifValue(OfficeExifTag.Artist, "Primary image");
            OfficeRasterEncodingResult encoded = OfficeRasterImageEncoder.EncodeWithMetadata(frames, OfficeImageExportFormat.Tiff, metadata);
            OfficeRasterFrames decoded = OfficeRasterImageDecoder.DecodeFrames(encoded.EncodedBytes, null, out OfficeRasterDecodeInfo info);
            Assert.True(info.DecodedAllFrames); Assert.False(info.FramesOrPagesDiscarded);
            Assert.Equal(OfficeImageFormat.Tiff, info.Format); Assert.Equal(2, info.Container!.Count);
            Assert.Equal(frames[0].Image.GetPixels(), decoded[0].Image.GetPixels());
            Assert.Equal(frames[1].Image.GetPixels(), decoded[1].Image.GetPixels());
            Assert.Equal(60D, OfficeImageMetadata.Read(encoded.EncodedBytes).Resolution.Horizontal);
            Assert.False(OfficeRasterImageDecoder.TryDecodeFrames(encoded.EncodedBytes, null, out _, out OfficeRasterDecodeInfo rejected, maximumFrames: 1));
            Assert.Equal(OfficeRasterDecodeFailure.FrameCountLimit, rejected.Failure);
            Assert.Equal(OfficeImageFormat.Tiff, rejected.Format);
            Assert.Throws<InvalidDataException>(() => OfficeRasterImageDecoder.DecodeFrames(encoded.EncodedBytes, maximumFrames: 1));
            Assert.Throws<NotSupportedException>(() => OfficeRasterImageEncoder.EncodeWithMetadata(frames, OfficeImageExportFormat.Png, metadata));
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        public void OrientationEvidenceDistinguishesStoredAndNormalizedSamples(OfficeImageExportFormat format) {
            byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), format);
            var metadata = OfficeImageMetadata.Read(bytes);
            metadata.SetExifValue(OfficeExifTag.Orientation, (ushort)6);
            bytes = OfficeImageMetadata.Apply(bytes, metadata);
            var source = new OfficeRasterDecodeOptions { ApplyExifOrientation = false };
            OfficeRasterDecodeOptions clone = source.Clone(); clone.ApplyExifOrientation = true;
            OfficeRasterFrames stored = OfficeRasterImageDecoder.DecodeFrames(bytes, source, out OfficeRasterDecodeInfo original);
            OfficeRasterFrames normalized = OfficeRasterImageDecoder.DecodeFrames(bytes, clone, out OfficeRasterDecodeInfo transformed);
            Assert.False(original.OrientationNormalized); Assert.True(transformed.OrientationNormalized);
            Assert.Equal((OfficeImageOrientation)6, transformed.SourceOrientation);
            Assert.Equal((OfficeImageOrientation)6, transformed.Container!.Frames[0].Orientation);
            Assert.Equal((3, 2), (stored[0].Image.Width, stored[0].Image.Height));
            Assert.Equal((2, 3), (normalized[0].Image.Width, normalized[0].Image.Height));
            Assert.False(source.ApplyExifOrientation);
            using var stream = new MemoryStream(bytes);
            OfficeRasterImageDecoder.DecodeFrames(stream, source, out OfficeRasterDecodeInfo streamed);
            Assert.Equal(0L, stream.Position); Assert.True(stream.CanRead); Assert.Equal(original.Format, streamed.Format);
        }

        [Fact]
        public void DiscoveryDistinguishesIdentificationOnlyAndManagedRasterRoutes() {
            foreach (OfficeImageFormat format in Enum.GetValues(typeof(OfficeImageFormat))) Assert.Equal(format, OfficeRasterImageFormats.GetCapabilities(format).Format);
            Assert.False(OfficeRasterImageFormats.GetCapabilities(OfficeImageFormat.Jpeg2000).CanDecode);
            Assert.True(OfficeRasterImageFormats.GetCapabilities(OfficeImageFormat.Jpeg2000).CanIdentify);
            Assert.False(OfficeRasterImageFormats.GetCapabilities(OfficeImageFormat.Webp).CanDecodeMultipleFrames);
            Assert.False(OfficeRasterImageFormats.GetCapabilities(OfficeImageFormat.Gif).CanEncode);
            foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Icon, OfficeImageExportFormat.Pbm, OfficeImageExportFormat.Tga }) {
                byte[] bytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue), format);
                Assert.Equal(format.GetContainerFormat(), OfficeRasterImageDecoder.DecodeFrames(bytes, null, out OfficeRasterDecodeInfo info).Count == 1 ? info.Format : OfficeImageFormat.Unknown);
                Assert.True(OfficeRasterImageFormats.GetCapabilities(info.Format).CanDecode);
            }
        }

        [Fact]
        public void MetadataEncodingReturnsNoPartialResultForLimitsOrCancellation() {
            var image = new OfficeRasterImage(3, 2, OfficeColor.Blue);
            var metadata = new OfficeImageMetadata { XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta>" + new string('a', 4096) + "</x:xmpmeta>") };
            Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Png, metadata, maximumEncodedBytes: 512));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeWithMetadata(image, OfficeImageExportFormat.Png, metadata, cancellationToken: cancellation.Token));
        }
        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        public void GenericExifResolutionEditsSurviveRealContainerRewrite(OfficeImageExportFormat format) {
            byte[] encoded = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), format);
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(encoded);
            metadata.SetExifValue(OfficeExifTag.XResolution, new OfficeRational(300, 1));
            metadata.SetExifValue(OfficeExifTag.YResolution, new OfficeRational(150, 1));
            metadata.SetExifValue(OfficeExifTag.ResolutionUnit, (ushort)2);

            OfficeImageMetadata rewritten = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(encoded, metadata));

            Assert.Equal(300D, rewritten.PhysicalDpiX);
            Assert.Equal(150D, rewritten.PhysicalDpiY);
            Assert.Equal(300D, Assert.IsType<OfficeRational>(rewritten.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
            Assert.Equal(150D, Assert.IsType<OfficeRational>(rewritten.GetExifValue(OfficeExifTag.YResolution)!.Value).ToDouble());
        }
    }
}
