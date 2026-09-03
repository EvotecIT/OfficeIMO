using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void OpcMetadataRewriteAppliesTheOutputLimitAfterCompression() {
        byte[] contentTypes = Encoding.UTF8.GetBytes(
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            new string(' ', 32 * 1024) +
            "<Override PartName=\"/META-INF/content_credential.c2pa\" ContentType=\"application/c2pa\"/>" +
            "</Types>");
        byte[] package = CreateCompressedZip(
            ("[Content_Types].xml", contentTypes),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = 128 * 1024;
        options.Limits.MaxManifestBytes = options.Limits.MaxAssetBytes;
        OfficeProvenanceRemovalResult baseline = OfficeProvenanceRemover.Remove(package, "document.docx", options);
        long finalSize = baseline.ToArray().LongLength;
        Assert.True(contentTypes.LongLength > finalSize);

        options.MaxOutputBytes = finalSize;
        OfficeProvenanceRemovalResult bounded = OfficeProvenanceRemover.Remove(package, "document.docx", options);

        Assert.Equal(finalSize, bounded.ToArray().LongLength);
        Assert.True(bounded.WasChanged);
    }

    [Fact]
    public void PackageMetadataReplacementAppliesTheOutputLimitAfterCompression() {
        byte[] replacement = Encoding.UTF8.GetBytes(
            "<metadata>" + new string(' ', 32 * 1024) + "</metadata>");
        byte[] package = CreateCompressedZip(
            ("remove.bin", new byte[] { 1 }),
            ("metadata.xml", Encoding.UTF8.GetBytes("<metadata/>")));
        OfficeProvenanceSignatureStripResult baseline = OfficeProvenanceZip.RemoveEntries(
            package,
            name => name == "remove.bin",
            maximumExpandedBytes: 128 * 1024,
            shouldReplace: name => name == "metadata.xml",
            replace: (_, _) => replacement,
            maximumReplacementBytes: 64 * 1024);
        Assert.True(replacement.LongLength > baseline.Data.LongLength);

        OfficeProvenanceSignatureStripResult bounded = OfficeProvenanceZip.RemoveEntries(
            package,
            name => name == "remove.bin",
            maximumExpandedBytes: 128 * 1024,
            shouldReplace: name => name == "metadata.xml",
            replace: (_, _) => replacement,
            maximumReplacementBytes: 64 * 1024,
            maximumOutputBytes: baseline.Data.LongLength);

        Assert.Equal(baseline.Data, bounded.Data);
    }
}
