using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FinalReplacementFailurePreservesInPlaceSourceAndRemovesStaging(bool exif) {
        // Windows range locks are mandatory. Unix locks do not reject a competing write.
        WithAtomicFiles((directory, source, _) => {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return;
            }
            byte[] original = File.ReadAllBytes(source);
            using (var locked = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                // Old input remains readable. The handle denies replacement, and this
                // mandatory range lock also rejects the expanded direct-write alternative.
                locked.Lock(original.Length + 8L, 1L);
                try {
                    Assert.ThrowsAny<IOException>(() => WriteAtomicProfile(source, source, exif));
                } finally {
                    locked.Unlock(original.Length + 8L, 1L);
                }
            }
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(new[] { source }, Directory.GetFiles(directory));
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FileWritersPublishCompleteNewAndExistingDestinations(bool exif, bool existing) {
        WithAtomicFiles((directory, source, target) => {
            byte[] original = File.ReadAllBytes(source);
            if (existing) {
                File.WriteAllBytes(target, new byte[] { 123, 45, 67 });
            }
            Assert.True(WriteAtomicProfile(source, target, exif));
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(target, out OfficeImageMetadata? profile));
            Assert.Equal(exif ? "Changed" : "Original", profile!.GetExifValue(OfficeExifTag.Software)!.Value);
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(target, out string? packet));
            Assert.Equal(exif ? "original-XMP" : "changed-XMP", packet);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(2, Directory.GetFiles(directory).Length);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancelledFileWriterPreservesExistingDestinationWithoutStaging(bool exif) {
        WithAtomicFiles((directory, source, target) => {
            byte[] original = File.ReadAllBytes(source);
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                WriteAtomicProfile(source, target, exif, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
            Assert.Equal(2, Directory.GetFiles(directory).Length);
        });
    }

    private static bool WriteAtomicProfile(string source, string target, bool exif, CancellationToken token = default) {
        if (!exif) {
            return OfficeHeifMetadataReader.TryWriteXmp(source, target, "changed-XMP", token);
        }
        var profile = new OfficeImageMetadata();
        profile.SetExifValue(OfficeExifTag.Software, "Changed");
        return OfficeHeifMetadataReader.TryWriteExifProfile(source, target, profile, token);
    }

    private static void WithAtomicFiles(Action<string, string, string> action) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-heif-atomic-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string source = Path.Combine(directory, "source.heic");
        try {
            File.WriteAllBytes(source, CoreHeifFixtures.CreateHeifMetadataSiblings(
                CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0));
            action(directory, source, Path.Combine(directory, "output.heic"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
