using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentDeferredPayloadPrerequisiteTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnderstatedDeflatedPayloadCannotPublishPrefixBytesOrACompleteHash(bool deferred) {
        byte[] bytes = SetZipField(Package(new byte[] { 1, 3, 5, 7 }), "Vendor/payload.bin", length: 3);
        if (!deferred) { Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(bytes))); return; }
        OdtDocument document = OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        Assert.Equal(3, Assert.Single(inspection.Payloads).Length);
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0));
        Assert.Throws<InvalidDataException>(() => document.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin"));
        Assert.Throws<InvalidDataException>(() => document.ToBytes());
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CorruptedZipChecksumCannotPublishPayloadBytesOrACompleteHash(bool deferred, bool storedControlPart) {
        byte[] bytes = SetZipField(Package(new byte[] { 1, 3, 5, 7 }), storedControlPart ? "mimetype" : "Vendor/payload.bin", corruptChecksum: true);
        if (storedControlPart) { Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(bytes),
            new OdfLoadOptions { DeferBinaryPayloads = deferred })); return; }
        if (!deferred) { Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(bytes))); return; }
        OdtDocument document = OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0));
        Assert.Throws<InvalidDataException>(() => document.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin"));
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
    }

    [Fact]
    public void DeferredSourceIsOwnedAndBoundedReadsDoNotMaterializeTheDocumentCache() {
        byte[] expected = { 1, 3, 5, 7 };
        byte[] bytes = Package(expected);
        using var source = new MemoryStream(bytes, writable: true);
        var options = new OdfLoadOptions { DeferBinaryPayloads = true };
        OdtDocument document = OdtDocument.Load(source, options);
        source.Dispose(); Array.Clear(bytes, 0, bytes.Length); options.DeferBinaryPayloads = false; options.MaxEntryUncompressedBytes = 1;
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
        Assert.Equal(4, document.GetPackageEntryByteLength("Vendor/payload.bin"));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin", 3));
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloads = 1, MaxPayloadBytes = 4, MaxTotalPayloadBytes = 4, MaxReadBytes = 4
        });
        OfficeEmbeddedPayloadInfo info = Assert.Single(inspection.Payloads);
        Assert.Equal(4, info.Length); Assert.Null(info.Sha256); Assert.Equal(0, inspection.ReadBytes);
        using Stream payload = inspection.OpenRead(info.Ordinal);
        inspection.Dispose(); Assert.Equal(expected, Read(payload));
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
        Assert.Equal(expected, document.GetPackageEntryBytes("Vendor/payload.bin", 4));
        Assert.Equal(OdfPackageEntryPayloadState.Loaded, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
    }

    [Fact]
    public void DeferredMetadataDoesNotExpandAnUnsupportedPayloadAndSelectedReadsStillFail() {
        byte[] bytes = SetZipField(Package(new byte[] { 1, 3, 5, 7 }), "Vendor/payload.bin", method: 99);
        Assert.Throws<InvalidDataException>(() => OdtDocument.Load(new MemoryStream(bytes)));
        OdtDocument document = OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { DeferBinaryPayloads = true });
        Assert.Equal("visible text", document.ContentBlocks[0].Paragraph!.Text);
        Assert.Equal(4, document.GetPackageEntryByteLength("Vendor/payload.bin"));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin", 3));
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        Assert.Null(Assert.Single(inspection.Payloads).Sha256);
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0));
        Assert.Equal(0, inspection.ReadBytes);
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
    }

    [Fact]
    public void DeferredLengthFailureCannotPublishAHashOrCacheIncompleteBytes() {
        byte[] bytes = SetZipField(Package(new byte[] { 1, 3, 5, 7 }), "Vendor/payload.bin", length: 5);
        OdtDocument document = OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0)); Assert.Equal(4, inspection.ReadBytes);
        Assert.Throws<InvalidDataException>(() => document.InspectEmbeddedPayloads(new OfficeEmbeddedPayloadInspectionOptions { IncludeSha256 = true }));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin"));
        Assert.Throws<InvalidDataException>(() => document.GetPackageEntryBytes("Vendor/payload.bin"));
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Vendor/payload.bin"));
    }

    [Fact]
    public void EditAndSavePreserveDeferredPayloadBytesAndOriginalSessionIdentity() {
        byte[] expected = { 1, 3, 5, 7 };
        OdtDocument document = OdtDocument.Load(new MemoryStream(Package(expected)), new OdfLoadOptions { DeferBinaryPayloads = true });
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        document.Metadata.Title = "Edited"; document.AddParagraph("new text");
        byte[] saved = document.ToBytes();
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(saved));
        Assert.Equal("Edited", reopened.Metadata.Title); Assert.Equal(2, reopened.ContentBlocks.Count);
        Assert.Equal(expected, reopened.GetPackageEntryBytes("Vendor/payload.bin"));
        using Stream payload = inspection.OpenRead(0); Assert.Equal(expected, Read(payload));
    }

    [Fact]
    public void DeferredEncryptedBinaryProfileIsRefusedWhileDefaultEncryptionRemainsVerified() {
        const string password = "test-odf-password";
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("secret");
        created.Package.AddOrReplaceEntry("Vendor/payload.bin", new byte[] { 1, 3, 5, 7 }, "application/octet-stream");
        byte[] bytes = created.ToBytes(new OdfSaveOptions { Encryption = new OdfEncryptionOptions { Password = password } });
        var options = new OdfLoadOptions { Password = password, DeferBinaryPayloads = true };
        OdfEncryptedPackageException refusal = Assert.Throws<OdfEncryptedPackageException>(() => OdtDocument.Load(new MemoryStream(bytes), options));
        Assert.Equal(OdfEncryptionFailureReason.UnsupportedProfile, refusal.Reason); Assert.Equal(password, options.Password);
        OdtDocument ordinary = OdtDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { Password = password });
        Assert.Equal(OdfPackageEntryPayloadState.Decrypted, ordinary.GetPackageEntryPayloadState("Vendor/payload.bin"));
        Assert.Equal(new byte[] { 1, 3, 5, 7 }, ordinary.GetPackageEntryBytes("Vendor/payload.bin"));
        options.MaxTotalKdfIterations = 1;
        Assert.Equal(OdfEncryptionFailureReason.ResourceLimitExceeded,
            Assert.Throws<OdfEncryptedPackageException>(() => OdtDocument.Load(new MemoryStream(bytes), options)).Reason);
    }

    [Theory]
    [InlineData("libreoffice-writer-basic.odt")]
    [InlineData("libreoffice-impress-basic.odp")]
    public void DeferredProducerPayloadMatchesTheIndependentArchiveMember(string fixture) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", fixture));
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        byte[] expected; using (Stream entry = archive.GetEntry("Thumbnails/thumbnail.png")!.Open()) expected = Read(entry);
        OdfDocument document = OdfDocument.Load(new MemoryStream(bytes), new OdfLoadOptions { DeferBinaryPayloads = true });
        Assert.Equal(OdfPackageEntryPayloadState.Deferred, document.GetPackageEntryPayloadState("Thumbnails/thumbnail.png"));
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        int ordinal = inspection.Payloads.Single(payload => payload.PartUri == "/Thumbnails/thumbnail.png").Ordinal;
        using Stream payload = inspection.OpenRead(ordinal); Assert.Equal(expected, Read(payload));
    }

    [Fact]
    public void DeferredEncryptedXmlProducerRetainsItsExistingPasswordAndIntegrityBoundary() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Encryption", "libreoffice-24.2-aes256.odt");
        var options = new OdfLoadOptions { DeferBinaryPayloads = true, Password = "OfficeIMO-ODF-test-2026" };
        OdtDocument document = OdtDocument.Load(path, options);
        Assert.Contains(document.ContentBlocks, block => block.Paragraph?.Text.Contains("LibreOffice encrypted producer fixture") == true);
        Assert.Equal(OdfPackageEntryPayloadState.Decrypted, document.GetPackageEntryPayloadState("content.xml"));
        options.Password = "wrong";
        Assert.Equal(OdfEncryptionFailureReason.IncorrectPassword,
            Assert.Throws<OdfEncryptedPackageException>(() => OdtDocument.Load(path, options)).Reason);
    }

    private static byte[] Package(byte[] payload) {
        OdtDocument document = OdtDocument.Create(); document.AddParagraph("visible text");
        return OdfTestPackageRewriter.Rewrite(document.ToBytes(), additions: new[] {
            new OdfTestPackageEntry("Vendor/payload.bin", payload)
        });
    }
    private static byte[] Read(Stream source) { using var output = new MemoryStream(); source.CopyTo(output); return output.ToArray(); }
    private static byte[] SetZipField(byte[] bytes, string target, ushort? method = null, uint? length = null, bool corruptChecksum = false) {
        // ZIP local and central header field offsets are independent fixture mutations, not owner helpers.
        for (int offset = 0; offset <= bytes.Length - 46; offset++) {
            uint signature = BitConverter.ToUInt32(bytes, offset);
            bool central = signature == 0x02014b50, local = signature == 0x04034b50;
            if (!central && !local) continue;
            int nameLength = BitConverter.ToUInt16(bytes, offset + (central ? 28 : 26));
            int nameOffset = offset + (central ? 46 : 30);
            if (nameOffset + nameLength > bytes.Length || Encoding.UTF8.GetString(bytes, nameOffset, nameLength) != target) continue;
            if (method.HasValue) Buffer.BlockCopy(BitConverter.GetBytes(method.Value), 0, bytes, offset + (central ? 10 : 8), 2);
            if (length.HasValue) Buffer.BlockCopy(BitConverter.GetBytes(length.Value), 0, bytes, offset + (central ? 24 : 22), 4);
            if (corruptChecksum) bytes[offset + (central ? 16 : 14)] ^= 1;
        }
        return bytes;
    }
}
