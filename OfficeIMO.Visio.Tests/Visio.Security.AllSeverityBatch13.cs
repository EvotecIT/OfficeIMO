using System.IO.Compression;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch13SecurityTests {
    [Fact]
    public void ValidatorRejectsOversizedEntryBeforeExtraction() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-vsdx-limit-" + Guid.NewGuid().ToString("N"));
        string package = Path.Combine(root, "oversized.vsdx");
        Directory.CreateDirectory(root);
        try {
            using (ZipArchive archive = ZipFile.Open(package, ZipArchiveMode.Create)) {
                ZipArchiveEntry entry = archive.CreateEntry("visio/document.xml", CompressionLevel.NoCompression);
                using Stream output = entry.Open();
                output.Write(new byte[32], 0, 32);
            }
            var validator = new VsdxPackageValidator(new VsdxPackageValidationLimits {
                MaxEntryBytes = 16,
                MaxTotalBytes = 64
            });

            Assert.False(validator.ValidateFile(package));
            Assert.Contains(validator.Errors, error => error.Contains("16-byte", StringComparison.Ordinal));
            Assert.False(validator.ValidateFileStreaming(package));
            Assert.Contains(validator.Errors, error => error.Contains("16-byte", StringComparison.Ordinal));

            string fixedPackage = Path.Combine(root, "fixed.vsdx");
            Assert.False(validator.FixFileStreaming(package, fixedPackage));
            Assert.Contains(validator.Errors, error => error.Contains("16-byte", StringComparison.Ordinal));
            Assert.False(File.Exists(fixedPackage));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void ValidatorEnforcesAggregateLimitAcrossDecompressedEntries() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-vsdx-aggregate-" + Guid.NewGuid().ToString("N"));
        string package = Path.Combine(root, "aggregate.vsdx");
        Directory.CreateDirectory(root);
        try {
            using (ZipArchive archive = ZipFile.Open(package, ZipArchiveMode.Create)) {
                for (int index = 0; index < 2; index++) {
                    ZipArchiveEntry entry = archive.CreateEntry("visio/entry" + index + ".bin", CompressionLevel.NoCompression);
                    using Stream output = entry.Open();
                    output.Write(new byte[12], 0, 12);
                }
            }
            var validator = new VsdxPackageValidator(new VsdxPackageValidationLimits {
                MaxEntryBytes = 16,
                MaxTotalBytes = 20
            });

            Assert.False(validator.ValidateFileStreaming(package));
            Assert.Contains(validator.Errors, error => error.Contains("20-byte aggregate", StringComparison.Ordinal));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void ValidatorRejectsTraversalEntryWithoutWritingOutsideTempRoot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-vsdx-traversal-" + Guid.NewGuid().ToString("N"));
        string package = Path.Combine(root, "traversal.vsdx");
        string escapedName = "escaped-" + Guid.NewGuid().ToString("N") + ".txt";
        string escapedPath = Path.Combine(Path.GetTempPath(), escapedName);
        Directory.CreateDirectory(root);
        try {
            using (ZipArchive archive = ZipFile.Open(package, ZipArchiveMode.Create)) {
                ZipArchiveEntry entry = archive.CreateEntry("../" + escapedName);
                using var writer = new StreamWriter(entry.Open());
                writer.Write("blocked");
            }
            var validator = new VsdxPackageValidator();

            Assert.False(validator.ValidateFile(package));
            Assert.Contains(validator.Errors, error => error.Contains("escapes", StringComparison.Ordinal));
            Assert.False(File.Exists(escapedPath));
        } finally {
            if (File.Exists(escapedPath)) File.Delete(escapedPath);
            Directory.Delete(root, true);
        }
    }
}
