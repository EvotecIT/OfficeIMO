using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void MarkdownFileLimitAppliesToPhysicalUtf16BytesRatherThanInternalUtf8Bytes() {
        string markdown = new string('\u0800', 2048);
        byte[] input = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(markdown)).ToArray();
        Assert.True(Encoding.UTF8.GetByteCount(markdown) > input.Length);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(path, input);
            var options = new OfficeProvenanceOptions {
                MaxAssetBytes = input.Length,
                MaxManifestBytes = Math.Min(1024, input.Length)
            };

            OfficeProvenanceReport report = MarkdownProvenance.InspectFile(path, options);

            Assert.Empty(report.Evidence);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
