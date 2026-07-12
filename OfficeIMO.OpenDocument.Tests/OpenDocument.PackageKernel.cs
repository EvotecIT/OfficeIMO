using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentPackageKernelTests {
    [Theory]
    [InlineData(OdfDocumentKind.Text, OdfMediaTypes.Text)]
    [InlineData(OdfDocumentKind.Spreadsheet, OdfMediaTypes.Spreadsheet)]
    [InlineData(OdfDocumentKind.Presentation, OdfMediaTypes.Presentation)]
    public void CreatesValidMinimalPackageWithRequiredMimetypeShape(OdfDocumentKind kind, string mediaType) {
        using OdfDocument document = Create(kind);

        byte[] bytes = document.ToBytes();

        Assert.True(document.Validate().IsValid);
        AssertMimetypeLocalHeader(bytes);
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.Equal("mimetype", archive.Entries[0].FullName);
        Assert.Equal(mediaType, ReadText(archive.GetEntry("mimetype")!));

        XDocument manifest = ReadXml(archive.GetEntry("META-INF/manifest.xml")!);
        XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        XElement rootEntry = manifest.Root!.Elements(ns + "file-entry")
            .Single(element => (string?)element.Attribute(ns + "full-path") == "/");
        Assert.Equal(mediaType, (string?)rootEntry.Attribute(ns + "media-type"));
        Assert.Equal("1.4", (string?)manifest.Root!.Attribute(ns + "version"));

        using OdfDocument reopened = Open(kind, bytes);
        Assert.Equal(kind, reopened.Kind);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void NativePackageWriterRoundTripsUnicodeEntryNames() {
        using OdtDocument document = OdtDocument.Create();
        byte[] expected = Encoding.UTF8.GetBytes("Zażółć gęślą jaźń");
        document.Package.AddOrReplaceEntry("Media/zażółć.txt", expected, "text/plain");

        byte[] bytes = document.ToBytes();

        using (var stream = new MemoryStream(bytes))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read)) {
            ZipArchiveEntry entry = archive.GetEntry("Media/zażółć.txt")!;
            using var output = new MemoryStream();
            using Stream input = entry.Open();
            input.CopyTo(output);
            Assert.Equal(expected, output.ToArray());
        }

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(bytes));
        Assert.Equal(expected, reopened.Package.GetRequiredEntry("Media/zażółć.txt").GetOriginalBytes());
    }

    [Fact]
    public void NativePackageWriterProducesStableDeterministicBytes() {
        using OdtDocument document = OdtDocument.Create();
        document.Metadata.Title = "Deterministic package";

        byte[] first = document.ToBytes();
        byte[] second = document.ToBytes();

        Assert.Equal(first, second);
    }

    [Fact]
    public void PreservesUnknownEntriesAndForeignXmlDuringTargetedMetadataEdit() {
        using OdtDocument created = OdtDocument.Create();
        created.Package.AddOrReplaceEntry("Vendor/custom.bin", new byte[] { 1, 3, 5, 7 }, "application/octet-stream");
        XDocument content = created.GetXml("content.xml");
        XNamespace vendor = "urn:example:vendor";
        content.Root!.Add(new XElement(vendor + "extension", new XAttribute(vendor + "flag", "keep")));
        created.MarkPartDirty("content.xml");
        byte[] source = created.ToBytes();

        using OdtDocument edited = OdtDocument.Open(new MemoryStream(source));
        edited.Metadata.Title = "Updated";
        OdfSaveResult save = edited.ToBytesResult();
        byte[] output = save.Value;

        Assert.Contains("content.xml", save.Report.CopiedEntries);
        Assert.Contains("meta.xml", save.Report.RewrittenEntries);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(output));
        Assert.Equal("Updated", reopened.Metadata.Title);
        Assert.Equal(new byte[] { 1, 3, 5, 7 }, reopened.Package.GetRequiredEntry("Vendor/custom.bin").GetOriginalBytes());
        Assert.NotNull(reopened.GetXml("content.xml").Root!.Element(vendor + "extension"));
    }

    [Fact]
    public void RejectsUnsafeArchiveEntryNames() {
        using OdtDocument document = OdtDocument.Create();
        byte[] valid = document.ToBytes();
        byte[] unsafePackage = OdfTestPackageRewriter.Rewrite(valid, additions: new[] {
            new OdfTestPackageEntry("../escape.xml", Encoding.UTF8.GetBytes("<escape/>"))
        });

        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(unsafePackage)));
    }

    [Fact]
    public void RejectsPackageBeyondConfiguredBudget() {
        using OdtDocument document = OdtDocument.Create();
        byte[] bytes = document.ToBytes();
        var options = new OdfOpenOptions { MaxPackageBytes = bytes.Length - 1 };

        Assert.Throws<InvalidDataException>(() => OdtDocument.Open(new MemoryStream(bytes), options));
    }

    [Fact]
    public void WritesOdf13CompatibilityProfileConsistently() {
        using OdsDocument document = OdsDocument.Create();
        byte[] bytes = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });

        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        Assert.Equal("1.3", (string?)ReadXml(archive.GetEntry("content.xml")!).Root!.Attribute(office + "version"));
        Assert.Equal("1.3", (string?)ReadXml(archive.GetEntry("META-INF/manifest.xml")!).Root!.Attribute(manifest + "version"));
    }

    [Fact]
    public void SerializationReportsDirtyStateUntilARealSaveSucceeds() {
        using OdtDocument document = OdtDocument.Create();
        using var destination = new MemoryStream();
        document.Save(destination);
        document.Metadata.Title = "Changed";

        OdfSaveResult firstSerialization = document.ToBytesResult();
        Assert.Contains("meta.xml", firstSerialization.Report.RewrittenEntries);

        OdfSaveResult repeatedSerialization = document.ToBytesResult();
        Assert.Contains("meta.xml", repeatedSerialization.Report.RewrittenEntries);

        document.Save(destination);
        OdfSaveResult acceptedSerialization = document.ToBytesResult();
        Assert.Empty(acceptedSerialization.Report.RewrittenEntries);
        Assert.Empty(acceptedSerialization.Report.RemovedEntries);
        Assert.Contains("meta.xml", acceptedSerialization.Report.CopiedEntries);
    }

    [Fact]
    public void FailedDestinationWriteDoesNotAcceptDirtyState() {
        using OdtDocument document = OdtDocument.Create();
        document.Save(new MemoryStream());
        document.Metadata.Title = "Pending";

        Assert.Throws<IOException>(() => document.Save(new ThrowingWriteStream()));
        OdfSaveResult serialization = document.ToBytesResult();

        Assert.Contains("meta.xml", serialization.Report.RewrittenEntries);
    }

    [Fact]
    public void SaveToStreamTruncatesAndRewindsDestination() {
        using OdtDocument document = OdtDocument.Create();
        using var destination = new MemoryStream(new byte[32_768], writable: true);

        document.Save(destination);

        Assert.Equal(0, destination.Position);
        Assert.True(destination.Length < 32_768);
        using OdtDocument reopened = OdtDocument.Open(destination);
        Assert.Equal(OdfDocumentKind.Text, reopened.Kind);
    }

    [Fact]
    public async Task SaveResultAsyncReturnsExactBytesAndOperationReport() {
        using OdtDocument document = OdtDocument.Create();
        document.Metadata.Title = "Async save result";
        using var destination = new MemoryStream();

        OdfSaveResult result = await document.SaveResultAsync(destination);

        Assert.Equal(result.Value, destination.ToArray());
        Assert.Same(result.Value, result.RequireNoLoss());
        Assert.Contains("meta.xml", result.Report.RewrittenEntries);
    }

    private static OdfDocument Create(OdfDocumentKind kind) {
        switch (kind) {
            case OdfDocumentKind.Text: return OdtDocument.Create();
            case OdfDocumentKind.Spreadsheet: return OdsDocument.Create();
            default: return OdpPresentation.Create();
        }
    }

    private static OdfDocument Open(OdfDocumentKind kind, byte[] bytes) {
        var stream = new MemoryStream(bytes);
        switch (kind) {
            case OdfDocumentKind.Text: return OdtDocument.Open(stream);
            case OdfDocumentKind.Spreadsheet: return OdsDocument.Open(stream);
            default: return OdpPresentation.Open(stream);
        }
    }

    private static void AssertMimetypeLocalHeader(byte[] bytes) {
        Assert.True(bytes.Length > 38);
        Assert.Equal(0x04034b50u, ReadUInt32(bytes, 0));
        Assert.Equal((ushort)0, ReadUInt16(bytes, 8));
        ushort nameLength = ReadUInt16(bytes, 26);
        ushort extraLength = ReadUInt16(bytes, 28);
        Assert.Equal("mimetype", Encoding.UTF8.GetString(bytes, 30, nameLength));
        Assert.Equal((ushort)0, extraLength);
    }

    private static string ReadText(ZipArchiveEntry entry) {
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static XDocument ReadXml(ZipArchiveEntry entry) {
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | (bytes[offset + 1] << 8));

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));

    private sealed class ThrowingWriteStream : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("Simulated destination failure.");
    }
}
