using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceReviewRegressionContracts {
    [Fact]
    public void JumbfStoreRejectsMalformedTrailingChildBox() {
        byte[] manifest = CreateManifestStore();
        Array.Resize(ref manifest, manifest.Length + 4);
        WriteBigEndian(manifest, 0, manifest.Length);
        byte[] png = CreatePngWithManifest(manifest);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void JumbfStoreRequiresAtLeastOneManifestSuperbox() {
        byte[] storeUuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] description = CreateBox("jumd", Join(storeUuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] png = CreatePngWithManifest(CreateBox("jumb", description));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void JumbfManifestSuperboxRequiresAClaimBox() {
        byte[] storeUuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] storeDescription = CreateBox("jumd", Join(storeUuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(manifestUuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("manifest\0")));
        byte[] descriptionOnlyManifest = CreateBox("jumb", manifestDescription);
        byte[] png = CreatePngWithManifest(CreateBox("jumb", Join(storeDescription, descriptionOnlyManifest)));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void PngChunkCountIsBoundedBeforeCarrierProcessing() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            png,
            "fixture.png",
            new OfficeProvenanceOptions { MaxContainerEntries = 1 }));
    }

    [Fact]
    public void MalformedOversizedTextWrapperIsRemovedAsOneCompleteRun() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        byte[] suffix = Encoding.UTF8.GetBytes("tail");
        byte[] text = Join(wrapper, suffix);
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };
        options.Limits.MaxManifestBytes = 1;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.txt", options);

        Assert.Equal(suffix, result.ToArray());
        Assert.Single(result.Changes);
    }

    [Fact]
    public void JpegInspectsAndRemovesAiDeclarationFromAdobeExtendedXmp() {
        byte[] extendedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        string guid = ComputeMd5(extendedPacket);
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" " +
            $"xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, extendedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport before = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");
        OfficeProvenanceReport after = OfficeProvenanceInspector.Inspect(result.ToArray(), "fixture.jpg");

        Assert.Single(before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(after.Evidence);
    }

    [Fact]
    public void JpegRejectsExtendedXmpWhosePacketDoesNotMatchTheReferencedDigest() {
        byte[] originalPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"/>");
        string guid = ComputeMd5(originalPacket);
        byte[] substitutedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, substitutedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Empty(report.Evidence);
        Assert.Contains(report.Diagnostics, item => item.Contains("digest", StringComparison.OrdinalIgnoreCase));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void XmpNodeBudgetAppliesBeforeStandardPacketMaterialization() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            string.Concat(Enumerable.Repeat("<x:n/>", 16)) +
            "<iptc:DigitalSourceType>trainedAlgorithmicMedia</iptc:DigitalSourceType></x:xmpmeta>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), packet)),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport bounded = OfficeProvenanceInspector.Inspect(
            jpeg, "fixture.jpg", new OfficeProvenanceOptions { MaxContainerEntries = 8 });
        OfficeProvenanceReport accepted = OfficeProvenanceInspector.Inspect(
            jpeg, "fixture.jpg", new OfficeProvenanceOptions { MaxContainerEntries = 64 });

        Assert.Empty(bounded.Evidence);
        Assert.Single(accepted.Evidence);
    }

    [Fact]
    public void SvgContentWinsOverGenericXmlFileExtension() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata><x iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></metadata></svg>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(svg, "fixture.xml");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.xml");

        Assert.Equal(OfficeProvenanceAssetFormat.Svg, report.Format);
        Assert.Single(report.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void Zip64EntryCountIsRejectedBeforeDirectoryMaterialization() {
        byte[] package = CreateZip64CountOnlyPackage(5000);

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            package,
            "fixture.zip",
            new OfficeProvenanceOptions { MaxContainerEntries = 10 }));
    }

    [Fact]
    public void ClassicZipMayContainExactlyTheSentinelEntryCountWithoutZip64Metadata() {
        byte[] endOfDirectory = new byte[22];
        WriteLittleEndian(endOfDirectory, 0, 0x06054B50U);
        WriteLittleEndian16(endOfDirectory, 8, ushort.MaxValue);
        WriteLittleEndian16(endOfDirectory, 10, ushort.MaxValue);

        OfficeProvenanceZip.ValidateEntryCount(endOfDirectory, ushort.MaxValue);
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceZip.ValidateEntryCount(endOfDirectory, ushort.MaxValue - 1));
    }

    [Fact]
    public void ZipRewritePreservesExternalAttributes() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                ZipArchiveEntry manifest = archive.CreateEntry("META-INF/content_credential.c2pa");
                using (Stream target = manifest.Open()) WriteAll(target, CreateManifestStore());
                ZipArchiveEntry script = archive.CreateEntry("bin/run.sh");
                script.ExternalAttributes = unchecked((int)0x81ED0000);
                using (Stream target = script.Open()) WriteAll(target, Encoding.UTF8.GetBytes("#!/bin/sh\n"));
            }
            package = AddCentralDirectoryComment(stream.ToArray(), "bin/run.sh", Encoding.UTF8.GetBytes("keep-comment"));
            package = AddArchiveComment(package, Encoding.UTF8.GetBytes("keep-archive-comment"));
        }

        Assert.Equal("keep-archive-comment", Encoding.UTF8.GetString(ReadArchiveComment(package)));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");
        using var rewritten = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.Equal(unchecked((int)0x81ED0000), rewritten.GetEntry("bin/run.sh")!.ExternalAttributes);
        int centralHeader = FindSignature(result.ToArray(), 0x02014B50u, "bin/run.sh");
        Assert.Equal(3, result.ToArray()[centralHeader + 5]);
        Assert.Equal("keep-comment", Encoding.UTF8.GetString(ReadCentralDirectoryComment(result.ToArray(), centralHeader)));
        Assert.Equal("keep-archive-comment", Encoding.UTF8.GetString(ReadArchiveComment(result.ToArray())));
    }

    [Fact]
    public void VerificationResultSnapshotsMutableFindings() {
        var findings = new List<string> { "initial" };
        var result = new OfficeProvenanceVerificationResult(
            OfficeProvenanceVerificationStatus.Valid, "test", findings);

        findings[0] = "changed";
        findings.Add("added");

        Assert.Equal(new[] { "initial" }, result.Findings);
    }

    [Fact]
    public void SignatureDiscoveryEnforcesAggregatePartBytesWhenDigestVerificationIsDisabled() {
        string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig2.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>";
        string signature =
            "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Object>" +
            new string('x', 2048) + "</Object></Signature>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(archive, "[Content_Types].xml", contentTypes);
                WriteZipEntry(archive, "_xmlsignatures/sig1.xml", signature);
                WriteZipEntry(archive, "_xmlsignatures/sig2.xml", signature);
            }
            package = output.ToArray();
        }
        int signatureBytes = Encoding.UTF8.GetByteCount(signature);
        var bounded = new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = signatureBytes + 1L,
            MaxTotalDigestBytes = signatureBytes + 1L
        };

        OfficePackageSignatureInfo rejected = OfficePackageSignatureService.Inspect(package, bounded);
        OfficePackageSignatureInfo accepted = OfficePackageSignatureService.Inspect(package,
            new OfficePackageSignatureInspectionOptions {
                VerifyDigests = false,
                MaxSignatureBytes = signatureBytes + 1L,
                MaxTotalDigestBytes = signatureBytes * 2L
            });

        Assert.Contains(rejected.SignatureParts, part =>
            part.ParseError?.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase) == true);
        Assert.DoesNotContain(accepted.SignatureParts, part =>
            part.ParseError?.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase) == true);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void IncompleteExtendedXmpDoesNotAllocateTheDeclaredPacketLength() {
        const string guid = "0123456789ABCDEF0123456789ABCDEF";
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" xmpNote:HasExtendedXMP=\"{guid}\"/>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, new byte[] { 1 }, 128 * 1024 * 1024),
            new byte[] { 0xFF, 0xD9 });

        long before = GC.GetAllocatedBytesForCurrentThread();
        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Empty(report.Evidence);
        Assert.True(allocated < 8L * 1024L * 1024L, $"Inspection allocated {allocated} bytes.");
    }


    [Fact]
    public void IncompleteApp11SequenceDoesNotAllocateTheDeclaredManifestLength() {
        byte[] fragment = CreateManifestStore();
        WriteBigEndian(fragment, 0, 64 * 1024 * 1024);
        byte[] app11Payload = Join(
            Encoding.ASCII.GetBytes("JP"),
            new byte[] { 0x12, 0x34 },
            BigEndian(1),
            fragment);
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xEB, app11Payload),
            new byte[] { 0xFF, 0xD9 });

        long before = GC.GetAllocatedBytesForCurrentThread();
        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.True(allocated < 8L * 1024L * 1024L, $"Inspection allocated {allocated} bytes.");
    }
#endif

    [Fact]
    public void ZipEmbeddedAssetsShareTheTopLevelCarrierLimit() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                using (Stream first = archive.CreateEntry("media/first.png").Open()) WriteAll(first, image);
                using (Stream second = archive.CreateEntry("media/second.png").Open()) WriteAll(second, image);
            }
            package = stream.ToArray();
        }
        var inspectionOptions = new OfficeProvenanceOptions { MaxCarriers = 1 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxCarriers = 1;

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(package, "fixture.zip", inspectionOptions));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(package, "fixture.zip", removalOptions));
    }

    [Fact]
    public void TiffIfdEntriesShareTheConfiguredContainerEntryLimit() {
        byte[] tiff = CreateTiffWithTwoIfds(entriesPerIfd: 2);
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 3 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(tiff, "fixture.tiff", options));

        Assert.Contains("container-entry limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SvgProcessesSeparateXmpAndDirectMetadataScopes() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" " +
            "xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<metadata><x:xmpmeta><rdf:RDF><rdf:Description/></rdf:RDF></x:xmpmeta></metadata>" +
            "<metadata><rdf:Description iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></metadata>" +
            "<rect width=\"1\" height=\"1\"/></svg>");

        OfficeProvenanceReport before = OfficeProvenanceInspector.Inspect(svg, "fixture.svg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Single(before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void SvgRejectsAValidDocumentBeforeMaterializingTooManyNodes() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><g/><g/><g/><g/><g/></svg>");
        var bounded = new OfficeProvenanceOptions { MaxContainerEntries = 5 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(svg, "fixture.svg", bounded));

        Assert.Contains("XML node limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(OfficeProvenanceInspector.Inspect(
            svg,
            "fixture.svg",
            new OfficeProvenanceOptions { MaxContainerEntries = 16 }).Evidence);
    }

    [Fact]
    public void RemovalResultSnapshotsMutableConstructorInputs() {
        byte[] data = { 1, 2, 3 };
        var changes = new List<OfficeProvenanceChange> {
            new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, "test", 1)
        };
        var report = new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Unknown, Array.Empty<OfficeProvenanceEvidence>());
        var result = new OfficeProvenanceRemovalResult(data, report, report, changes, wasReserialized: false);

        data[0] = 9;
        changes.Clear();

        Assert.Equal(new byte[] { 1, 2, 3 }, result.ToArray());
        Assert.True(result.WasChanged);
        Assert.Single(result.Changes);
    }

    private static void WriteAll(Stream stream, byte[] data) => stream.Write(data, 0, data.Length);

    private static void WriteZipEntry(ZipArchive archive, string name, string content) {
        using Stream stream = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
        WriteAll(stream, Encoding.UTF8.GetBytes(content));
    }

    private static byte[] CreateManifestStore() {
        byte[] uuid = { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] descriptionPayload = Join(uuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa\0"));
        byte[] description = CreateBox("jumd", descriptionPayload);
        byte[] manifestUuid = { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] manifestDescription = CreateBox("jumd", Join(manifestUuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] claimUuid = { 0x63, 0x32, 0x63, 0x6C, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 };
        byte[] claimDescription = CreateBox("jumd", Join(claimUuid, new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        return CreateBox("jumb", Join(description, CreateBox("jumb", Join(manifestDescription, claim))));
    }

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[8 + payload.Length];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type, 0, 4, box, 4);
        Buffer.BlockCopy(payload, 0, box, 8, payload.Length);
        return box;
    }

    private static byte[] CreatePngWithManifest(byte[] manifest) => Join(
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
        CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
        CreatePngChunk("caBX", manifest),
        CreatePngChunk("IEND", Array.Empty<byte>()));

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        byte[] chunk = new byte[12 + payload.Length];
        WriteBigEndian(chunk, 0, payload.Length);
        Buffer.BlockCopy(typeBytes, 0, chunk, 4, 4);
        Buffer.BlockCopy(payload, 0, chunk, 8, payload.Length);
        WriteBigEndian(chunk, 8 + payload.Length, unchecked((int)Crc32(Join(typeBytes, payload))));
        return chunk;
    }

    private static uint Crc32(byte[] data) {
        uint crc = 0xFFFFFFFF;
        foreach (byte value in data) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ (0xEDB88320U & (uint)-(int)(crc & 1));
        }
        return ~crc;
    }

    private static byte[] CreateTextWrapper(byte[] manifest) {
        byte[] header = Join(Encoding.ASCII.GetBytes("C2PATXT\0"), new byte[] { 1 }, BigEndian(manifest.Length), manifest);
        var builder = new StringBuilder("\uFEFF");
        foreach (byte value in header) builder.Append(char.ConvertFromUtf32(value < 16 ? 0xFE00 + value : 0xE0100 + value - 16));
        return Encoding.UTF8.GetBytes(builder.ToString());
    }

    private static byte[] CreateExtendedXmpSegment(string guid, byte[] packet) {
        return CreateExtendedXmpSegment(guid, packet, packet.Length);
    }

    private static byte[] CreateExtendedXmpSegment(string guid, byte[] packet, int declaredLength) {
        byte[] header = Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
        byte[] payload = new byte[header.Length + 40 + packet.Length];
        Buffer.BlockCopy(header, 0, payload, 0, header.Length);
        Encoding.ASCII.GetBytes(guid, 0, guid.Length, payload, header.Length);
        WriteBigEndian(payload, header.Length + 32, declaredLength);
        WriteBigEndian(payload, header.Length + 36, 0);
        Buffer.BlockCopy(packet, 0, payload, header.Length + 40, packet.Length);
        return CreateJpegSegment(0xE1, payload);
    }

    private static int FindSignature(byte[] data, uint signature, string entryName) {
        byte[] name = Encoding.UTF8.GetBytes(entryName);
        for (int index = 0; index + 46 + name.Length <= data.Length; index++) {
            if (BitConverter.ToUInt32(data, index) != signature) continue;
            int nameLength = BitConverter.ToUInt16(data, index + 28);
            if (nameLength == name.Length && data.AsSpan(index + 46, nameLength).SequenceEqual(name)) return index;
        }
        throw new InvalidDataException("ZIP central-directory entry was not found.");
    }

    private static byte[] AddCentralDirectoryComment(byte[] package, string entryName, byte[] comment) {
        int centralHeader = FindSignature(package, 0x02014B50u, entryName);
        int nameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        int extraLength = BitConverter.ToUInt16(package, centralHeader + 30);
        Assert.Equal(0, BitConverter.ToUInt16(package, centralHeader + 32));
        int insertOffset = centralHeader + 46 + nameLength + extraLength;
        int endOffset = -1;
        for (int index = package.Length - 22; index >= 0; index--) {
            if (BitConverter.ToUInt32(package, index) == 0x06054B50u) { endOffset = index; break; }
        }
        if (endOffset < 0) throw new InvalidDataException("ZIP end record was not found.");
        byte[] updated = new byte[package.Length + comment.Length];
        Buffer.BlockCopy(package, 0, updated, 0, insertOffset);
        Buffer.BlockCopy(comment, 0, updated, insertOffset, comment.Length);
        Buffer.BlockCopy(package, insertOffset, updated, insertOffset + comment.Length, package.Length - insertOffset);
        WriteLittleEndian16(updated, centralHeader + 32, checked((ushort)comment.Length));
        int updatedEndOffset = endOffset + comment.Length;
        uint centralSize = BitConverter.ToUInt32(updated, updatedEndOffset + 12);
        WriteLittleEndian(updated, updatedEndOffset + 12, checked(centralSize + (uint)comment.Length));
        return updated;
    }

    private static byte[] ReadCentralDirectoryComment(byte[] package, int centralHeader) {
        int nameLength = BitConverter.ToUInt16(package, centralHeader + 28);
        int extraLength = BitConverter.ToUInt16(package, centralHeader + 30);
        int commentLength = BitConverter.ToUInt16(package, centralHeader + 32);
        byte[] comment = new byte[commentLength];
        Buffer.BlockCopy(package, centralHeader + 46 + nameLength + extraLength, comment, 0, commentLength);
        return comment;
    }

    private static byte[] AddArchiveComment(byte[] package, byte[] comment) {
        int endOffset = FindEndOfCentralDirectory(package);
        Assert.Equal(0, BitConverter.ToUInt16(package, endOffset + 20));
        byte[] updated = new byte[package.Length + comment.Length];
        Buffer.BlockCopy(package, 0, updated, 0, package.Length);
        WriteLittleEndian16(updated, endOffset + 20, checked((ushort)comment.Length));
        Buffer.BlockCopy(comment, 0, updated, package.Length, comment.Length);
        return updated;
    }

    private static byte[] ReadArchiveComment(byte[] package) {
        int endOffset = FindEndOfCentralDirectory(package);
        int length = BitConverter.ToUInt16(package, endOffset + 20);
        return package.AsSpan(endOffset + 22, length).ToArray();
    }

    private static int FindEndOfCentralDirectory(byte[] package) {
        for (int index = package.Length - 22; index >= Math.Max(0, package.Length - 22 - ushort.MaxValue); index--) {
            if (BitConverter.ToUInt32(package, index) == 0x06054B50u &&
                index + 22 + BitConverter.ToUInt16(package, index + 20) == package.Length) return index;
        }
        throw new InvalidDataException("ZIP end record was not found.");
    }

    private static byte[] CreateJpegSegment(byte marker, byte[] payload) {
        byte[] segment = new byte[payload.Length + 4];
        segment[0] = 0xFF;
        segment[1] = marker;
        int length = payload.Length + 2;
        segment[2] = (byte)(length >> 8);
        segment[3] = (byte)length;
        Buffer.BlockCopy(payload, 0, segment, 4, payload.Length);
        return segment;
    }

    private static string ComputeMd5(byte[] data) {
        using MD5 md5 = MD5.Create();
        return string.Concat(md5.ComputeHash(data).Select(value => value.ToString("X2")));
    }

    private static byte[] CreateZip64CountOnlyPackage(ulong count) {
        byte[] package = new byte[102];
        package[0] = 0x50; package[1] = 0x4B; package[2] = 0x03; package[3] = 0x04;
        WriteLittleEndian(package, 4, 0x06064B50U);
        WriteLittleEndian64(package, 8, 44);
        WriteLittleEndian64(package, 28, count);
        WriteLittleEndian64(package, 36, count);
        WriteLittleEndian(package, 60, 0x07064B50U);
        WriteLittleEndian64(package, 68, 4);
        WriteLittleEndian(package, 76, 1U);
        WriteLittleEndian(package, 80, 0x06054B50U);
        package[88] = 0xFF; package[89] = 0xFF;
        package[90] = 0xFF; package[91] = 0xFF;
        return package;
    }

    private static byte[] CreateTiffWithTwoIfds(int entriesPerIfd) {
        int ifdSize = 2 + entriesPerIfd * 12 + 4;
        byte[] data = new byte[8 + ifdSize * 2];
        data[0] = (byte)'I';
        data[1] = (byte)'I';
        data[2] = 42;
        WriteLittleEndian(data, 4, 8U);
        WriteLittleEndian16(data, 8, (ushort)entriesPerIfd);
        WriteLittleEndian(data, 8 + 2 + entriesPerIfd * 12, (uint)(8 + ifdSize));
        WriteLittleEndian16(data, 8 + ifdSize, (ushort)entriesPerIfd);
        return data;
    }

    private static byte[] BigEndian(int value) {
        byte[] bytes = new byte[4];
        WriteBigEndian(bytes, 0, value);
        return bytes;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static void WriteLittleEndian(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteLittleEndian16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteLittleEndian64(byte[] data, int offset, ulong value) {
        for (int index = 0; index < 8; index++) data[offset + index] = (byte)(value >> (index * 8));
    }

    private static byte[] Join(params byte[][] values) {
        byte[] result = new byte[values.Sum(value => value.Length)];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, result, offset, value.Length);
            offset += value.Length;
        }
        return result;
    }
}
