using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave78Tests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InspectionAndRemovalHonorUtf32Preambles(bool bigEndian) {
        string input = Path.Combine(Path.GetTempPath(), "officeimo-html-utf32-" + Guid.NewGuid().ToString("N") + ".html");
        string output = Path.Combine(Path.GetTempPath(), "officeimo-html-utf32-" + Guid.NewGuid().ToString("N") + ".html");
        var encoding = new UTF32Encoding(bigEndian, byteOrderMark: true);
        const string html = "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>utf32</body></html>";
        File.WriteAllBytes(input, encoding.GetPreamble().Concat(encoding.GetBytes(html)).ToArray());

        try {
            OfficeProvenanceReport before = HtmlProvenance.InspectFile(input);
            OfficeProvenanceRemovalResult removal = HtmlProvenance.RemoveFile(input, output);
            OfficeProvenanceReport after = HtmlProvenance.InspectFile(output);
            byte[] published = File.ReadAllBytes(output);

            Assert.Equal(
                OfficeProvenanceCarrierKind.C2paExternalManifest,
                Assert.Single(before.Evidence).Carrier);
            Assert.True(removal.WasChanged);
            Assert.Empty(after.Evidence);
            Assert.True(encoding.GetPreamble().SequenceEqual(published.Take(encoding.GetPreamble().Length)));
        } finally {
            File.Delete(input);
            File.Delete(output);
        }
    }
}
