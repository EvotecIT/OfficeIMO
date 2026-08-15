using System.IO.Compression;
using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlIgnoresInactivePictureSources() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<picture><source media=\"print\" type=\"image/png\" srcset=\"{dataUri}\"><img src=\"fallback.png\"></picture>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlIgnoresPictureFallbackWhenAnEarlierSourceIsSelected() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<picture><source type=\"image/png\" srcset=\"selected.png\"><img src=\"{dataUri}\"></picture>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlIgnoresImageDeclarationsForUnmatchedSelectors() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<style>.unused{{background-image:url({dataUri})}}</style><div class=\"used\"></div>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlPreservesSrcdocLocationForEmbeddedImages() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string nested = $"<img src=\"{dataUri}\">".Replace("\"", "&quot;");
        string html = $"<iframe srcdoc=\"{nested}\"></iframe>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);

        Assert.StartsWith("HTML/iframe[srcdoc][0]/img[src][0]", Assert.Single(report.Evidence).Location, StringComparison.Ordinal);
    }

    [Fact]
    public void EpubApiRejectsForeignSignedZipBeforeMutation() {
        byte[] package = CreateEpubTestZip(
            ("mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")),
            ("META-INF/signatures.xml", Encoding.UTF8.GetBytes("<document-signatures/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package));
        Assert.Throws<InvalidDataException>(() => {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".epub");
            try { File.WriteAllBytes(path, package); EpubDocument.InspectProvenance(path); }
            finally { File.Delete(path); }
        });
    }

    private static byte[] CreateEpubTestZip(params (string Name, byte[] Data)[] entries) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] data) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                using Stream target = entry.Open();
                target.Write(data, 0, data.Length);
            }
        }
        return output.ToArray();
    }
}
