using System;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed partial class DrawingHeifMetadataTests {
        [Theory]
        [InlineData(OfficeImageMetadataProfileKinds.Exif, false)]
        [InlineData(OfficeImageMetadataProfileKinds.Xmp, false)]
        [InlineData(OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, false)]
        [InlineData(OfficeImageMetadataProfileKinds.Exif, true)]
        [InlineData(OfficeImageMetadataProfileKinds.Xmp, true)]
        [InlineData(OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, true)]
        public void CombinedProfileWritesReplaceOrClearOnlySelectedFamilies(OfficeImageMetadataProfileKinds profiles, bool clear) {
            byte[] bytes = CoreHeifFixtures.CreateMinimalHeifWithPrimaryImageExifAndXmp(320, 180,
                CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP");
            byte[] original = (byte[])bytes.Clone();
            OfficeImageMetadata? metadata = clear ? null : new OfficeImageMetadata();
            if (metadata != null) {
                metadata.SetExifValue(OfficeExifTag.Software, "Changed");
                metadata.XmpProfile = Encoding.UTF8.GetBytes("<xmp>日本語 \U0001F600</xmp>");
            }
            Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? before));
            Assert.True(OfficeHeifMetadataReader.TryWriteMetadata(bytes, metadata, profiles, out byte[]? output));
            Assert.NotSame(bytes, output);
            Assert.Equal(original, bytes);
            Assert.True(OfficeHeifMetadataReader.TryReadInfo(output!, out OfficeHeifImageInfo? after));
            Assert.Equal(before!.Width, after!.Width);
            Assert.Equal(before.Height, after.Height);
            bool editExif = (profiles & OfficeImageMetadataProfileKinds.Exif) != 0;
            bool editXmp = (profiles & OfficeImageMetadataProfileKinds.Xmp) != 0;
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(output!, out OfficeImageMetadata? exif));
            Assert.Equal(editExif && clear ? null : editExif ? "Changed" : "Original", exif?.GetExifValue(OfficeExifTag.Software)?.Value);
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(output!, out string? xmp));
            Assert.Equal(editXmp && clear ? string.Empty : editXmp ? "<xmp>日本語 \U0001F600</xmp>" : "original-XMP", xmp);
            if (!editExif) {
                Assert.Equal(ProfileBytes(bytes, before.ExifItem!), ProfileBytes(output!, after.ExifItem!));
            }
            if (!editXmp) {
                Assert.Equal(ProfileBytes(bytes, before.XmpItem!), ProfileBytes(output!, after.XmpItem!));
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        public void CombinedSecondEditFailureNeverReturnsFirstEditBytes(int rejection) {
            byte[] bytes = rejection == 6 ? CoreHeifFixtures.CreateMinimalHeifWithExif(CoreHeifFixtures.CreateExifPayload("Original"))
                : CoreHeifFixtures.CreateHeifMetadataSiblings(CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0,
                    rejection <= 3 ? rejection : 0, encoding: rejection == 5 ? "gzip" : "", xmpProtection: (ushort)(rejection == 4 ? 1 : 0));
            byte[] original = (byte[])bytes.Clone();
            OfficeImageMetadata metadata = new OfficeImageMetadata();
            metadata.SetExifValue(OfficeExifTag.Software, "Changed");
            metadata.XmpProfile = rejection == 7 ? new byte[] { 0x80 } : Encoding.UTF8.GetBytes("changed-XMP");
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out _));
            Assert.False(OfficeHeifMetadataReader.TryWriteMetadata(bytes, metadata,
                OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, out byte[]? output));
            Assert.Null(output);
            Assert.Equal(original, bytes);
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out OfficeImageMetadata? exif));
            Assert.Equal("Original", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CombinedMissingProfilesClearWithoutSynthesizingResolutionExif(bool replaceXmp) {
            byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0);
            OfficeImageMetadata metadata = new OfficeImageMetadata {
                ExifProfile = null, Resolution = new OfficeImageResolution(144.5D, 82D),
                XmpProfile = replaceXmp ? Encoding.UTF8.GetBytes("changed-XMP") : null
            };
            Assert.False(metadata.HasExifProfile);
            Assert.True(OfficeHeifMetadataReader.TryWriteMetadata(bytes, metadata,
                OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, out byte[]? output));
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(output!, out OfficeImageMetadata? exif));
            Assert.Null(exif);
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(output!, out string? xmp));
            Assert.Equal(replaceXmp ? "changed-XMP" : string.Empty, xmp);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CombinedSelectiveExifEditKeepsUnsupportedOrMalformedXmpOpaque(bool encoded) {
            byte[] packet = encoded ? Encoding.UTF8.GetBytes("opaque") : new byte[] { 0x80 };
            byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(CoreHeifFixtures.CreateExifPayload("Original"), "unused", 0, 0,
                encoding: encoded ? "gzip" : "", xmpBytes: packet);
            OfficeImageMetadata metadata = new OfficeImageMetadata();
            metadata.SetExifValue(OfficeExifTag.Software, "Changed");
            metadata.XmpProfile = new byte[] { 0x80 };
            Assert.True(OfficeHeifMetadataReader.TryWriteMetadata(bytes, metadata, OfficeImageMetadataProfileKinds.Exif, out byte[]? output));
            Assert.True(OfficeHeifMetadataReader.TryReadInfo(output!, out OfficeHeifImageInfo? info));
            Assert.Equal(packet, ProfileBytes(output!, info!.XmpItem!));
        }

        [Fact]
        public void CombinedNoOpCopiesMetadataLessContainersAndHonorsCancellation() {
            byte[] bytes = CoreHeifFixtures.CreateMinimalHeifWithoutExif();
            byte[] original = (byte[])bytes.Clone();
            Assert.True(OfficeHeifMetadataReader.TryWriteMetadata(bytes, null, OfficeImageMetadataProfileKinds.None, out byte[]? output));
            Assert.NotSame(bytes, output);
            Assert.Equal(bytes, output);
            using CancellationTokenSource cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            output = new byte[] { 123 };
            Assert.Throws<OperationCanceledException>(() => OfficeHeifMetadataReader.TryWriteMetadata(bytes, null,
                OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, out output, cancellation.Token));
            Assert.Null(output);
            Assert.Equal(original, bytes);
        }

        [Theory]
        [InlineData(OfficeImageMetadataProfileKinds.Icc)]
        [InlineData((OfficeImageMetadataProfileKinds)64)]
        public void CombinedWritesRejectUnsupportedProfileMasks(OfficeImageMetadataProfileKinds profiles) {
            byte[] bytes = CoreHeifFixtures.CreateMinimalHeifWithoutExif();
            byte[]? output = new byte[] { 123 };
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeHeifMetadataReader.TryWriteMetadata(bytes, null, profiles, out output));
            Assert.Null(output);
        }

        [Fact]
        public void CombinedClearChargesOriginalInputAndRetainedFirstOutput() {
            byte[] bytes = CreateLargeHeifSource(exif: false, prefix: 0);
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out _));
            Assert.True(OfficeHeifMetadataReader.TryWriteXmp(bytes, null, out _));
            Assert.False(OfficeHeifMetadataReader.TryWriteMetadata(bytes, null,
                OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp, out byte[]? output));
            Assert.Null(output);
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? original));
            Assert.Equal(6 * 1024 * 1024, original!.Length);
        }

        [Fact]
        public void CombinedSelectiveWriteChargesCallerOwnedUnselectedProfiles() {
            byte[] bytes = CreateLargeHeifSource(exif: false, prefix: 0);
            OfficeImageMetadata metadata = new OfficeImageMetadata {
                XmpProfile = new byte[16 * 1024 * 1024], IccProfile = new byte[16 * 1024 * 1024]
            };
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out _));
            Assert.False(OfficeHeifMetadataReader.TryWriteMetadata(bytes, metadata, OfficeImageMetadataProfileKinds.Exif, out byte[]? output));
            Assert.Null(output);
        }

        private static byte[] ProfileBytes(byte[] data, OfficeHeifItemInfo item) {
            OfficeHeifItemExtentInfo extent = item.Location!.Extents.Single();
            return data.Skip(extent.Offset).Take(extent.Length).ToArray();
        }
    }
}
