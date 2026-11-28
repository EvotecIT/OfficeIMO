using System.Buffers.Binary;
using System.IO.Compression;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointEmbeddedPayloadTests {
    private const string Slide = "/ppt/slides/slide1.xml";
    private const string PackagePart = "/ppt/embeddings/attached%20file.bin";
    private const string PayloadId = "rIdEmbeddedPackage";
    private const string PayloadIdentity = Slide + "#" + PayloadId;
    private const string RelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package";

    [Fact]
    public void OwnedInventoryKeepsTheSdkDescriptorAndIgnoresAnOrphanMentionedInAnExtension() {
        byte[] package = CreateFixture(128);
        using var baseline = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        _ = baseline.OpenXmlDocument;
        OfficeEmbeddedPayloadInfo expected = Assert.Single(baseline.GetEmbeddedPayloads(includeSha256: true));
        using var presentation = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        OfficeEmbeddedPayloadInfo actual = Assert.Single(presentation.GetEmbeddedPayloads(includeSha256: true));
        Assert.Equal(OfficeEmbeddedPayloadKind.EmbeddedPackage, actual.Kind);
        Assert.Equal(PayloadIdentity, actual.Id);
        Assert.Equal(expected.PartUri, actual.PartUri);
        Assert.Equal(expected.ContentType, actual.ContentType);
        Assert.Equal(expected.SuggestedFileName, actual.SuggestedFileName);
        Assert.Equal(expected.Length, actual.Length);
        Assert.Equal(expected.Sha256, actual.Sha256);
        Assert.False(presentation.IsSdkLoaded);

        presentation.Slides[0].AddTextBox("Pending slide edit");
        OfficeEmbeddedPayloadInfo metadata = Assert.Single(presentation.GetEmbeddedPayloads());
        Assert.Equal(actual.Length, metadata.Length);
        Assert.Null(metadata.Sha256);
        Assert.False(presentation.IsSdkLoaded);
    }

    [Fact]
    public void OwnedReadsRespectDecodedWorkLimitsAndReturnDetachedStreams() {
        const int payloadLength = 2 * 1024 * 1024;
        byte[] package = CreateFixture(payloadLength, incompressible: true);
        using var source = new CountingInput(package);
        using var presentation = PowerPointDocument.Load(source, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        using var bounded = presentation.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloadBytes = payloadLength, MaxTotalPayloadBytes = payloadLength, MaxReadBytes = 64
        });
        OfficeEmbeddedPayloadInfo metadata = Assert.Single(bounded.Payloads);
        Assert.False(metadata.HasKnownLength);
        Assert.Equal(0, bounded.ReadBytes);

        source.BytesRead = 0;
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => bounded.OpenRead(0));
        Assert.Equal("MaxReadBytes", exception.LimitName);
        Assert.Equal(65, bounded.ReadBytes);
        Assert.InRange(source.BytesRead, 1, 64 * 1024);

        using Stream detached = presentation.OpenEmbeddedPayload(PayloadIdentity, payloadLength);
        Assert.False(presentation.IsSdkLoaded);
        using var session = presentation.InspectEmbeddedPayloads();
        _ = presentation.OpenXmlDocument;
        using Stream rebound = session.OpenRead(0);
        Assert.Equal(payloadLength, rebound.Length);
        presentation.Dispose();

        Assert.Equal(payloadLength, detached.Length);
        Assert.Equal(29, detached.ReadByte());
    }

    [Fact]
    public void InterleavedRelationshipsStillExposeTheOriginalEmbeddedPayload() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateInterleavedFixture(), writable: false));
        OfficeEmbeddedPayloadInfo info = Assert.Single(presentation.GetEmbeddedPayloads());
        Assert.Equal(PayloadIdentity, info.Id);
        Assert.Equal(256, info.Length);
    }

    [Fact]
    public void ClearingSlidesRemovesAssetsOwnedByAnInterleavedRelationshipPart() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateInterleavedFixture(), writable: false));
        Assert.Single(presentation.Slides);

        presentation.ClearSlides();
        byte[] saved = presentation.ToBytes();

        using var package = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
        Assert.False(package.PartExists(Slide));
        Assert.False(package.PartExists(PackagePart));
        Assert.False(package.HasInterleavedParts);
    }

    private static byte[] CreateInterleavedFixture() {
        byte[] fixture = CreateFixture(256);
        using var bytes = new MemoryStream();
        bytes.Write(fixture, 0, fixture.Length);
        bytes.Position = 0;
        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Update, leaveOpen: true)) {
            string path = OfficeOpcPartNames.GetRelationshipsPartName(Slide).TrimStart('/');
            ZipArchiveEntry original = archive.GetEntry(path)!;
            using var xml = new MemoryStream();
            using (Stream input = original.Open()) input.CopyTo(xml);
            original.Delete();
            byte[] content = xml.ToArray();
            int middle = content.Length / 2;
            using (Stream first = archive.CreateEntry(path + "/[0].piece").Open())
                first.Write(content, 0, middle);
            using (Stream last = archive.CreateEntry(path + "/[1].last.piece").Open())
                last.Write(content, middle, content.Length - middle);
        }

        return bytes.ToArray();
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MalformedRelationshipsDoNotProducePayloadDescriptors(bool duplicateId) {
        using var storage = OfficeOpcEditablePackage.Open(new MemoryStream(CreateFixture(256), writable: false));
        storage.CreateBinaryPart("/ppt/embeddings/second.bin", "application/octet-stream", new byte[] { 2 });
        if (duplicateId) {
            const int ns = OfficeXmlNamespaces.PackageRelationships;
            OfficeXmlTree relationships = storage.GetRelationships(Slide);
            int relationship = relationships.CreateElement(ns, "Relationship");
            relationships.SetAttribute(relationship, OfficeXmlNamespaces.None, "Id", PayloadId);
            relationships.SetAttribute(relationship, OfficeXmlNamespaces.None, "Type", RelationshipType);
            relationships.SetAttribute(relationship, OfficeXmlNamespaces.None, "Target", "/ppt/embeddings/second.bin");
            relationships.Append(relationships.Root, relationship);
        } else {
            storage.AddRelationship(Slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideLayout", "/ppt/embeddings/second.bin");
        }
        using var bytes = new MemoryStream();
        storage.Save(bytes);
        byte[] fixture = bytes.ToArray();
        using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        Exception? expected = Record.Exception(() => {
            _ = baseline.OpenXmlDocument;
            baseline.InspectEmbeddedPayloads().Dispose();
        });

        Assert.NotNull(expected);
        Assert.IsType(expected.GetType(), Record.Exception(() => presentation.InspectEmbeddedPayloads().Dispose()));
    }

    [Fact]
    public void PayloadReadsAndHashesStopAtTheDeclaredDecodedLength() {
        byte[] fixture = CreateFixture(256);
        SetDeclaredPayloadSize(fixture, 1);
        using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        _ = baseline.OpenXmlDocument;
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        OfficeEmbeddedPayloadInfo metadata = Assert.Single(presentation.GetEmbeddedPayloads());
        Assert.Equal(1, metadata.Length);
#if NETFRAMEWORK
        Exception? baselineFailure = Record.Exception(() => baseline.GetEmbeddedPayloads(includeSha256: true));
        Assert.Equal("System.IO.FileFormatException", baselineFailure?.GetType().FullName);
        Assert.Throws<InvalidDataException>(() => presentation.GetEmbeddedPayloads(includeSha256: true));
        Assert.Throws<InvalidDataException>(() => presentation.OpenEmbeddedPayload(PayloadIdentity, 1));
#else
        OfficeEmbeddedPayloadInfo expected = Assert.Single(baseline.GetEmbeddedPayloads(includeSha256: true));
        OfficeEmbeddedPayloadInfo hashed = Assert.Single(presentation.GetEmbeddedPayloads(includeSha256: true));
        using Stream payload = presentation.OpenEmbeddedPayload(PayloadIdentity, 1);
        Assert.Equal(expected.Length, hashed.Length);
        Assert.Equal(expected.Sha256, hashed.Sha256);
        Assert.Equal(1, payload.Length);
        Assert.Equal(29, payload.ReadByte());
#endif
        Assert.False(presentation.IsSdkLoaded);
    }

    private static void SetDeclaredPayloadSize(byte[] fixture, uint size) {
        int end = fixture.Length - 22; // These fixtures have neither a ZIP comment nor a ZIP64 directory.
        Assert.Equal(0x06054b50u, BinaryPrimitives.ReadUInt32LittleEndian(fixture.AsSpan(end)));
        int entry = BinaryPrimitives.ReadInt32LittleEndian(fixture.AsSpan(end + 16));
        while (entry < end) {
            Assert.Equal(0x02014b50u, BinaryPrimitives.ReadUInt32LittleEndian(fixture.AsSpan(entry)));
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(fixture.AsSpan(entry + 28));
            int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(fixture.AsSpan(entry + 30));
            int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(fixture.AsSpan(entry + 32));
            if (Encoding.UTF8.GetString(fixture, entry + 46, nameLength) == PackagePart.TrimStart('/')) {
                Assert.Equal(8, BinaryPrimitives.ReadUInt16LittleEndian(fixture.AsSpan(entry + 10)));
                int local = BinaryPrimitives.ReadInt32LittleEndian(fixture.AsSpan(entry + 42));
                BinaryPrimitives.WriteUInt32LittleEndian(fixture.AsSpan(local + 22), size);
                BinaryPrimitives.WriteUInt32LittleEndian(fixture.AsSpan(entry + 24), size);
                return;
            }
            entry += 46 + nameLength + extraLength + commentLength;
        }
        throw new InvalidOperationException("The fixture is missing its payload.");
    }

    private sealed class CountingInput(byte[] data) : MemoryStream(data, writable: false) {
        internal long BytesRead { get; set; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
    }

    private static byte[] CreateFixture(int length, bool incompressible = false) {
        using var presentation = PowerPointDocument.Create();
        presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        using var package = OfficeOpcEditablePackage.Open(new MemoryStream(presentation.ToBytes(), writable: false));
        byte[] data = new byte[length];
        if (incompressible) new Random(3706).NextBytes(data);
        else for (int index = 0; index < data.Length; index++) data[index] = 29;
        data[0] = 29;
        package.CreateBinaryPart(PackagePart, "application/octet-stream", data);
        package.AddRelationship(Slide, RelationshipType, PackagePart, id: PayloadId);
        package.CreateBinaryPart("/ppt/embeddings/orphan.bin", "application/octet-stream", new byte[] { 1 });
        OfficeXmlTree xml = package.GetXml(Slide);
        int list = xml.CreateElement(OfficeXmlNamespaces.PresentationML, "extLst");
        xml.Append(xml.Root, list);
        int extension = xml.CreateElement(OfficeXmlNamespaces.PresentationML, "ext");
        xml.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "urn:fixture:unknown");
        xml.Append(list, extension);
        int decoy = xml.CreateElement(OfficeXmlNamespaces.PresentationML, "oleObj");
        xml.SetAttribute(decoy, OfficeXmlNamespaces.Relationships, "id", "rIdOrphan");
        xml.Append(extension, decoy);
        using var output = new MemoryStream();
        package.Save(output);
        return output.ToArray();
    }
}
