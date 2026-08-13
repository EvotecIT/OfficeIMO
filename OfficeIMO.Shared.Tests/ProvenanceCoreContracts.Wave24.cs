using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ManifestRejectsRawDirectChildrenButAcceptsExtensionSuperboxes() {
        byte[] rawChild = CreateBox("cbor", new byte[] { 0xA0 });
        byte[] extensionDescription = CreateBox("jumd", Join(
            C2paUuid("priv"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("vendor.example\0")));
        byte[] extensionChild = CreateBox("jumb", Join(extensionDescription, CreateBox("cbor", new byte[] { 0xA0 })));

        OfficeProvenanceReport raw = OfficeProvenanceInspector.Inspect(
            CreatePngWithC2paManifest(AppendDirectManifestChild(CreateManifestStore(), rawChild)),
            "fixture.png");
        OfficeProvenanceReport extension = OfficeProvenanceInspector.Inspect(
            CreatePngWithC2paManifest(AppendDirectManifestChild(CreateManifestStore(), extensionChild)),
            "fixture.png");

        Assert.False(Assert.Single(raw.Evidence).IsStructurallyValid);
        Assert.True(Assert.Single(extension.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ManifestAcceptsTheLegacyOptionalDataBoxStore() {
        byte[] description = CreateBox("jumd", Join(
            C2paUuid("c2db"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.databoxes\0")));
        byte[] dataBoxes = CreateBox("jumb", Join(description, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] png = CreatePngWithC2paManifest(AppendDirectManifestChild(CreateManifestStore(), dataBoxes));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ManifestRejectsMalformedLegacyDataBoxStoreContent() {
        byte[] description = CreateBox("jumd", Join(
            C2paUuid("c2db"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.databoxes\0")));
        byte[] dataBoxes = CreateBox("jumb", Join(description, CreateBox("free", new byte[] { 0x00 })));
        byte[] png = CreatePngWithC2paManifest(AppendDirectManifestChild(CreateManifestStore(), dataBoxes));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void OpcMetadataCleanupHonorsTheConfiguredXmlNodeBudget() {
        string overrides = string.Concat(Enumerable.Range(0, 32).Select(index =>
            $"<Override PartName=\"/word/media/image{index}.png\" ContentType=\"image/png\"/>"));
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                overrides +
                "<Override PartName=\"/META-INF/content_credential.c2pa\" ContentType=\"application/c2pa\"/></Types>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(package, "document.docx", options));
    }

    [Fact]
    public void OpcReferencesRemainWhenAnyDuplicateNativeManifestIsPreserved() {
        const string relationship =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"urn:c2pa\" Target=\"META-INF/content_credential.c2pa\"/>" +
            "</Relationships>";
        const string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/META-INF/content_credential.c2pa\" ContentType=\"application/c2pa\"/>" +
            "</Types>";
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("_rels/.rels", Encoding.UTF8.GetBytes(relationship)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("META-INF/content_credential.c2pa", new byte[] { 1, 2, 3 }));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "document.docx", new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.PreserveSignatureMarkup
        });
        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.Single(archive.Entries, entry => entry.FullName == "META-INF/content_credential.c2pa");
        Assert.Contains("content_credential.c2pa", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "_rels/.rels")), StringComparison.Ordinal);
        Assert.Contains("content_credential.c2pa", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "[Content_Types].xml")), StringComparison.Ordinal);
    }

    [Fact]
    public void Gif87aC2paApplicationExtensionIsStructurallyInvalid() {
        byte[] application = CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore());
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF87a"), new byte[7], application, new byte[] { 0x3B });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(gif, "fixture.gif");
        OfficeProvenanceRemovalResult strict = OfficeProvenanceRemover.Remove(gif, "fixture.gif");
        OfficeProvenanceRemovalResult relaxed = OfficeProvenanceRemover.Remove(gif, "fixture.gif", new OfficeProvenanceRemovalOptions {
            RequireStructurallyValidCarrier = false
        });

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(strict.WasChanged);
        Assert.True(relaxed.WasChanged);
        Assert.Empty(relaxed.After.Evidence);
    }

    private static byte[] AppendDirectManifestChild(byte[] store, byte[] child) {
        int storeDescriptionLength = ReadBigEndianInt32(store, 8);
        int manifestOffset = 8 + storeDescriptionLength;
        int manifestLength = ReadBigEndianInt32(store, manifestOffset);
        byte[] result = new byte[store.Length + child.Length];
        Buffer.BlockCopy(store, 0, result, 0, store.Length);
        Buffer.BlockCopy(child, 0, result, store.Length, child.Length);
        WriteBigEndian(result, manifestOffset, manifestLength + child.Length);
        WriteBigEndian(result, 0, result.Length);
        return result;
    }

    private static int ReadBigEndianInt32(byte[] data, int offset) =>
        data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3];

    private static byte[] CreatePngWithC2paManifest(byte[] manifest) => Join(
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
        CreatePngChunk("IHDR", new byte[13]),
        CreatePngChunk("caBX", manifest),
        CreatePngChunk("IEND", Array.Empty<byte>()));
}
