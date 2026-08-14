using System.Text;
using System.IO.Compression;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void PreservedTextWrapperStillMakesExternalManifestAmbiguous() {
        byte[] external = Encoding.UTF8.GetBytes(
            "-----BEGIN C2PA MANIFEST-----\nhttps://example.test/manifest.c2pa\n-----END C2PA MANIFEST-----\n");
        byte[] input = Join(external, CreateTextWrapper(CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = false,
            RemoveExternalC2paReferences = true
        };

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt", options);

        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void OrphanExtendedJpegXmpMakesTheStandardPacketAmbiguous() {
        byte[] standard = CreateJpegSegment(
            0xE1,
            Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), CreateXmpPacket()));
        byte[] extended = CreateJpegSegment(
            0xE1,
            Join(
                Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0"),
                Encoding.ASCII.GetBytes("0123456789ABCDEF0123456789ABCDEF"),
                new byte[] { 0, 0, 0, 1, 0, 0, 0, 0, (byte)'x' }));
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            standard,
            extended,
            CreateMinimalJpegFrame(),
            CreateMinimalJpegScan(),
            new byte[] { 0, 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void GifWithoutAnImageDescriptorPreservesC2pa() {
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[7],
            CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore()),
            new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void JpegWithoutAFrameAndScanPreservesC2pa() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void WebpWithoutAnExtendedHeaderPreservesC2pa() {
        byte[] webp = CreateWebp(
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("C2PA", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void PngWithInvalidHeaderFieldsPreservesC2pa() {
        byte[] invalidHeader = { 0, 0, 0, 0, 0, 0, 0, 1, 8, 2, 0, 0, 0 };
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", invalidHeader),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", Array.Empty<byte>()),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData("<htmlish>")]
    [InlineData("<!doctype html-not-really>")]
    public void HtmlLikePrefixesDoNotHideStructuredTextCarriers(string prefix) {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        byte[] input = Encoding.UTF8.GetBytes(
            prefix + "\n-----BEGIN C2PA MANIFEST-----\ndata:application/c2pa;base64," + manifest +
            "\n-----END C2PA MANIFEST-----\n");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(input, "fixture.txt");

        Assert.Equal(OfficeProvenanceAssetFormat.StructuredText, report.Format);
        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void OwningNonOpcPackageCanPreserveUnrelatedContentTypesMetadata() {
        byte[] package = CreateZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes("not-opc-xml")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.RemoveZipPackage(
            package,
            "fixture.odt",
            new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.PreserveSignatureMarkup
            },
            removeOpcManifestReferences: false);

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.Null(archive.GetEntry("META-INF/content_credential.c2pa"));
        Assert.Equal("not-opc-xml", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "[Content_Types].xml")));
    }

    private static byte[] CreateMinimalJpegFrame() => CreateJpegSegment(
        0xC0,
        new byte[] { 8, 0, 1, 0, 1, 1, 1, 0x11, 0 });

    private static byte[] CreateMinimalJpegScan() => CreateJpegSegment(
        0xDA,
        new byte[] { 1, 1, 0, 0, 63, 0 });

    private static byte[] CreateValidPngHeader() =>
        new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 };

    private static byte[] CreateMinimalGifImage() =>
        new byte[] { 0x2C, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0, 0 };
}
