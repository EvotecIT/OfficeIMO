using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void HtmlCommentFormatDetectionHonorsTheContainerEntryLimit() {
        byte[] text = Encoding.UTF8.GetBytes("<!--a--b--><!--c--d-->plain");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(
                text,
                "fixture.txt",
                new OfficeProvenanceOptions { MaxContainerEntries = 1 }));

        Assert.Contains("HTML format detection", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WrapperPrefixDetectionHonorsTheContainerEntryLimit() {
        byte[] text = Encoding.UTF8.GetBytes("\uFEFF\uFEFF\uFEFFplain");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(
                text,
                "fixture.bin",
                new OfficeProvenanceOptions { MaxContainerEntries = 2 }));

        Assert.Contains("wrapper detection", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void JumbfManifestLabelMustBeValidUtf8() {
        byte[] manifest = CreateManifestStore();
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        int uuidOffset = FindSequenceOffset(manifest, manifestUuid);
        Assert.True(uuidOffset >= 0);
        manifest[uuidOffset + manifestUuid.Length + 1] = 0xFF;
        byte[] png = CreatePngWithManifest(manifest);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData(".png")]
    [InlineData(".svg")]
    public void ZipPreservesStructuredTextMislabeledAsAnImage(string extension) {
        string carrier = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                using Stream target = archive.CreateEntry("media/not-an-image" + extension, CompressionLevel.Optimal).Open();
                WriteAll(target, Encoding.UTF8.GetBytes(carrier));
            }
            package = output.ToArray();
        }

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "fixture.zip");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(package, result.ToArray());
    }

    private static int FindSequenceOffset(byte[] data, byte[] expected) {
        for (int offset = 0; offset <= data.Length - expected.Length; offset++) {
            bool matches = true;
            for (int index = 0; index < expected.Length; index++) {
                if (data[offset + index] == expected[index]) continue;
                matches = false;
                break;
            }
            if (matches) return offset;
        }
        return -1;
    }
}
