using System.IO.Compression;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch17SecurityTests {
    [Fact]
    public void ValidatorRejectsUnixLinkMetadataAcrossTargetFrameworks() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-vsdx-link-" + Guid.NewGuid().ToString("N") + ".vsdx");
        try {
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Create)) {
                ZipArchiveEntry entry = archive.CreateEntry("visio/document.xml");
                typeof(ZipArchiveEntry).GetProperty("ExternalAttributes")!
                    .SetValue(entry, unchecked((int)0xA0000000), null);
                using StreamWriter writer = new StreamWriter(entry.Open());
                writer.Write("<document />");
            }

            var validator = new VsdxPackageValidator();

            Assert.False(validator.ValidateFile(path));
            Assert.Contains(validator.Errors, error =>
                error.Contains("link entry", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
