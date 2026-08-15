using System.IO.Compression;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ProvenanceSupportsKeepsActiveImagePropertiesInScope() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style>@supports (mask-image:none){.target{mask-image:url('" +
            dataUri + "')}}</style></head><body><div class='target'></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(report.HasC2paManifest);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData("0x")]
    [InlineData("-0x")]
    [InlineData("0.0x")]
    public void SrcsetRejectsNonPositiveDensityDescriptors(string descriptor) {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string srcset = dataUri + " " + descriptor;

        Assert.Empty(HtmlSrcSetParser.Parse(srcset));
        Assert.Empty(HtmlProvenance.Inspect("<img srcset=\"" + srcset + "\">").Evidence);
    }

    [Theory]
    [InlineData("Pictures/../image.png")]
    [InlineData("./Pictures/image.png")]
    public void OdfProvenanceRejectsUnsafeOrNonCanonicalEntryPaths(string entryName) {
        byte[] package = CreateZipPackage(
            "odt",
            "META-INF/documentsignatures.xml",
            CreatePngWithManifest(CreateManifestStore()));
        using var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            WriteEntry(archive, entryName, new byte[] { 1, 2, 3 }, CompressionLevel.Optimal);
        }
        byte[] unsafePackage = output.ToArray();
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".odt");
        try {
            File.WriteAllBytes(path, unsafePackage);
            Assert.Throws<InvalidDataException>(() => OdfDocument.InspectProvenance(path));
            Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(unsafePackage, "document.odt"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
