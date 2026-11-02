using System;
using System.IO;
using System.Text;
using OfficeIMO.Reader.Internal.IO;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderMaterializationAdmissionTests {
    [Theory]
    [InlineData("media")]
    [InlineData("table")]
    [InlineData("visual")]
    public void AnOwnedDestinationIsSkippedWithoutLeavingStagingFiles(string family) {
        string directory = NewDirectory();
        try {
            string path = OutputPath(directory, family);
            using (OfficeFileCommitClaim claim = OfficeFileCommitClaim.Acquire(path)) {
                Assert.False(Materialize(directory, family, "contending payload"));
                Assert.False(File.Exists(path));
            }
            Assert.Empty(Directory.GetFileSystemEntries(directory));
            Assert.True(Materialize(directory, family, "complete payload"));
            Assert.Equal("complete payload", File.ReadAllText(path));
        } finally { Directory.Delete(directory, true); }
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void AnInaccessibleAbandonedClaimRemainsAPermissionFailure() {
        if (OperatingSystem.IsWindows() || Environment.UserName == "root") return;
        string directory = NewDirectory();
        string claim = OfficeFileCommitClaim.CreatePath(OutputPath(directory, "media"));
        try {
            File.WriteAllText(claim, "abandoned");
            File.SetUnixFileMode(claim, UnixFileMode.None);
            Assert.ThrowsAny<UnauthorizedAccessException>(() => Materialize(directory, "media", "payload"));
            Assert.False(File.Exists(OutputPath(directory, "media")));
        } finally {
            if (File.Exists(claim)) File.SetUnixFileMode(claim, UnixFileMode.UserRead | UnixFileMode.UserWrite);
            Directory.Delete(directory, true);
        }
    }
#endif

    [Fact]
    public void AnUnrelatedDestinationFailureIsNotReportedAsSuccessfulOrSkipped() {
        string directory = NewDirectory();
        try {
            string destination = OutputPath(directory, "media");
            Directory.CreateDirectory(destination);
            Assert.ThrowsAny<IOException>(() => Materialize(directory, "media", "payload"));
            Assert.True(Directory.Exists(destination));
            Assert.Equal(new[] { destination }, Directory.GetFileSystemEntries(directory));
        } finally { Directory.Delete(directory, true); }
    }

    private static bool Materialize(string directory, string family, string payload) {
        if (family == "media") {
            var document = new OfficeDocumentReadResult {
                Assets = new[] { new OfficeDocumentAsset { Id = "shared", FileName = "shared.bin", PayloadBytes = Encoding.UTF8.GetBytes(payload) } }
            };
            return Assert.Single(document.WriteAssetsToDirectory(directory, new OfficeDocumentAssetMaterializationOptions { Overwrite = false })).Written;
        }
        if (family == "table") {
            var exports = new[] { new ReaderTableExportBundle { Id = "shared", FileNamePrefix = "shared", Csv = payload } };
            return Assert.Single(exports.WriteTableExportsToDirectory(directory, new ReaderTableExportMaterializationOptions {
                Overwrite = false, IncludeCsv = true, IncludeMarkdown = false, IncludeJson = false
            })).Written;
        }
        var visuals = new[] { new ReaderVisualExportBundle { Id = "shared", FileNamePrefix = "shared", PayloadExtension = ".txt", Payload = payload } };
        return Assert.Single(visuals.WriteVisualExportsToDirectory(directory, new ReaderVisualExportMaterializationOptions {
            Overwrite = false, IncludePayload = true, IncludeJson = false
        })).Written;
    }

    private static string OutputPath(string directory, string family) => Path.Combine(directory,
        family == "media" ? "shared.bin" : family == "table" ? "shared.csv" : "shared.txt");

    private static string NewDirectory() {
        string path = Path.Combine(Path.GetTempPath(), "reader-materialization-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path); return path;
    }
}
