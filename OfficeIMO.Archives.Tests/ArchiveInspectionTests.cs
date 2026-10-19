using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveInspectionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicMemberReadUsesTheExistingMsZipDecoderForACompleteFrame(bool grouped) {
        // CAB CK marker followed by a final stored-DEFLATE block containing the independent text "hello".
        byte[] frame = { 0x43, 0x4b, 1, 5, 0, 0xfa, 0xff, (byte)'h', (byte)'e', (byte)'l', (byte)'l', (byte)'o' };
        byte[] cabinet = ArchiveFixtures.StoredCab("message.txt", frame);
        cabinet[42] = 1;
        Buffer.BlockCopy(BitConverter.GetBytes(5u), 0, cabinet, 44, 4);
        int dataOffset = BitConverter.ToInt32(cabinet, 36);
        cabinet[dataOffset + 6] = 5;
        using var source = new MemoryStream(cabinet);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.Equal(ArchiveCompression.MsZip, Assert.Single(inspection.Entries).Compression);
        using ArchiveReadBatch? batch = grouped ? inspection.ReadMembers(new[] { 0 }, 5) : null;
        using Stream member = batch == null ? inspection.OpenRead(0) : batch.Members[0].Content;
        Assert.Equal("hello", Text(member));
    }

    [Fact]
    public void EmptyZipInspectionPreservesTheCallerPositionAndOwnership() {
        using var source = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(source, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true)) { }
        source.Position = 3;
        using var inspection = ArchiveInspector.Open(source);
        Assert.Equal(ArchiveFormat.Zip, inspection.Format);
        Assert.Empty(inspection.Entries);
        Assert.Equal(3, source.Position);
        Assert.True(source.CanRead);
    }

    [Fact]
    public void RawCabNamesAndDuplicateRecordsRemainSeparateWithoutFilesystemExtraction() {
        byte[] cabinet = ArchiveFixtures.StoredCab(("../escape.txt", Encoding.UTF8.GetBytes("first")),
            ("../escape.txt", Encoding.UTF8.GetBytes("second")));
        using var source = new MemoryStream(cabinet);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.Equal(ArchiveFormat.Cab, inspection.Format);
        Assert.Empty(inspection.StorageStreams);
        Assert.Equal(new[] { "../escape.txt", "../escape.txt" }, inspection.Entries.Select(entry => entry.RawName));
        Assert.Equal(new[] { 0, 1 }, inspection.Entries.Select(entry => entry.Ordinal));
        Assert.Equal(new long[] { 5, 6 }, inspection.Entries.Select(entry => entry.Length));
        Assert.All(inspection.Entries, entry => {
            Assert.Null(entry.AuthoredPath); Assert.True(entry.CanRead);
            Assert.Equal(ArchiveCompression.Store, entry.Compression);
            Assert.Equal(11, entry.FolderExpandedBytes); Assert.Equal(11, entry.FolderCompressedBytes);
            Assert.Equal(1, entry.FolderFrameCount); Assert.Equal(0, entry.FolderChecksumFrameCount);
        });
        using Stream first = inspection.OpenRead(0);
        using Stream second = inspection.OpenRead(1);
        Assert.Equal("first", Text(first)); Assert.Equal("second", Text(second));
        Assert.False(first.CanWrite); Assert.False(second.CanWrite);
    }

    [Theory]
    [InlineData("makecab-lzx-testOneNote2016.cab", "testOneNote2016.one")]
    [InlineData("makecab-lzx-testOneNoteFromOffice365-2.cab", "testOneNoteFromOffice365-2.one")]
    public void MicrosoftMakecabPayloadMatchesTheIndependentOriginalThroughPublicApi(string cabinet, string original) {
        string fixtures = Path.Combine(AppContext.BaseDirectory, "Fixtures");
        byte[] expected = File.ReadAllBytes(Path.Combine(fixtures, original));
        using ArchiveInspection inspection = ArchiveInspector.Open(Path.Combine(fixtures, cabinet));
        ArchiveInspectionEntry entry = Assert.Single(inspection.Entries);
        Assert.Equal(original, entry.RawName); Assert.Equal(expected.LongLength, entry.Length);
        Assert.Equal(ArchiveCompression.Lzx, entry.Compression); Assert.True(entry.CanRead);
        Assert.True(entry.FolderChecksumFrameCount > 0);
        using Stream payload = inspection.OpenRead(entry.Ordinal);
        Assert.Equal(expected, Bytes(payload));
    }

    [Fact]
    public void MsiInventoryKeepsFileRowsAndReadsBothMappedPayloadsAndRawStorageStreams() {
        using var source = new MemoryStream(ArchiveFixtures.Msi(twoCabinets: true));
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.Equal(ArchiveFormat.Msi, inspection.Format);
        Assert.Equal(new[] { "PayloadKey", "PayloadTwo" }, inspection.Entries.Select(entry => entry.RawName));
        Assert.Equal("TEMPLA~1.ADM|Template.adml", inspection.Entries[0].RawFileName);
        Assert.Equal("PolicyDefinitions/en-US/Template.adml", inspection.Entries[0].AuthoredPath);
        Assert.Equal("#payload.cab", inspection.Entries[0].CabinetName);
        Assert.Equal(1, inspection.Entries[0].Sequence); Assert.Equal(0, inspection.Entries[0].CabinetOrdinal);
        using Stream first = inspection.OpenRead(0); using Stream second = inspection.OpenRead(1);
        Assert.Equal("hello", Text(first)); Assert.Equal("world", Text(second));
        ArchiveStorageStream cabinet = Assert.Single(inspection.StorageStreams, stream =>
            stream.RawName == MsiPackageReader.EncodeStreamName("payload.cab", false));
        using Stream raw = inspection.OpenStorageStream(cabinet.Ordinal);
        Assert.Equal(ArchiveFixtures.StoredCab("PayloadKey", Encoding.UTF8.GetBytes("hello")), Bytes(raw));
        Assert.Equal(inspection.StorageStreams.OrderBy(stream => stream.DirectoryOrder).Select(stream => stream.RawName),
            inspection.StorageStreams.Select(stream => stream.RawName));
    }

    [Fact]
    public void MsiInspectionPreservesAuthoredSpellingWhileExtractionUsesInstalledNames() {
        const string directory = "Templates. /PolicyDefinitions";
        const string filename = "SHORT~1|Template.admx . ";
        byte[] bytes = ArchiveFixtures.MsiAliasFiles(false, new MsiFixtureFile(filename, "A", parent: directory));
        using var source = new MemoryStream(bytes);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        ArchiveInspectionEntry entry = Assert.Single(inspection.Entries);
        Assert.Equal(filename, entry.RawFileName);
        Assert.Equal(directory + "/Template.admx . ", entry.AuthoredPath);
        using Stream payload = inspection.OpenRead(entry.Ordinal);
        Assert.Equal("A", Text(payload));
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult extracted = ArchiveExtractor.Extract(fixture.Package(bytes), fixture.Destination);
        Assert.Equal("Templates/PolicyDefinitions/Template.admx", Assert.Single(extracted.Files).RelativePath);
        Assert.Equal("A", File.ReadAllText(Path.Combine(fixture.Destination, "Templates", "PolicyDefinitions", "Template.admx")));
    }

    [Fact]
    public void MsiAliasesAreInventoriedIndependentlyEvenWhenTheirBytesDiffer() {
        byte[] msi = ArchiveFixtures.Msi(twoCabinets: true, editDatabase: (_, rows) => rows["File"][1][2] = "Template.adml");
        using var source = new MemoryStream(msi);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.Equal(2, inspection.Entries.Count);
        Assert.Equal(inspection.Entries[0].AuthoredPath, inspection.Entries[1].AuthoredPath);
        using Stream first = inspection.OpenRead(0); using Stream second = inspection.OpenRead(1);
        Assert.Equal("hello", Text(first)); Assert.Equal("world", Text(second));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExternalCabinetsAndLooseSourcesAreRetainedWithoutOpeningTheirPaths(bool external) {
        byte[] msi = ArchiveFixtures.Msi(externalCab: external, editDatabase: (_, rows) => {
            if (external) rows["Media"][0][3] = "https://example.invalid/do-not-fetch.cab";
            else rows["File"][0][6] = 0x2000;
        });
        using var source = new MemoryStream(msi);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        ArchiveInspectionEntry entry = Assert.Single(inspection.Entries);
        Assert.Equal("PayloadKey", entry.RawName); Assert.Equal(5, entry.Length); Assert.False(entry.CanRead);
        Assert.Equal(external ? "MSI_EXTERNAL_CABINET" : "MSI_LOOSE_SOURCE", entry.Diagnostic!.Code);
        if (external) Assert.Equal("https://example.invalid/do-not-fetch.cab", entry.CabinetName);
        Assert.Throws<NotSupportedException>(() => inspection.OpenRead(0));
    }

    [Theory]
    [InlineData(2, ArchiveCompression.Quantum, "CAB_COMPRESSION")]
    [InlineData(0x0e03, ArchiveCompression.Lzx, "CAB_LZX_WINDOW")]
    public void UnsupportedFolderCompressionRemainsVisibleAndCannotReturnPayload(int rawCompression,
        ArchiveCompression compression, string code) {
        byte[] cabinet = ArchiveFixtures.StoredCab("file", new byte[] { 1 });
        cabinet[42] = (byte)rawCompression; cabinet[43] = (byte)(rawCompression >> 8);
        using var source = new MemoryStream(cabinet);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        ArchiveInspectionEntry entry = Assert.Single(inspection.Entries);
        Assert.Equal(compression, entry.Compression); Assert.Equal((ushort)rawCompression, entry.RawCompression);
        Assert.False(entry.CanRead); Assert.Equal(code, entry.Diagnostic!.Code);
        Assert.StartsWith(code + ":", Assert.Throws<NotSupportedException>(() => inspection.OpenRead(0)).Message);
        Assert.Equal(0, inspection.ReadBytes);
    }

    [Fact]
    public void InventoryDoesNotClaimPayloadVerificationAndCorruptChecksumsNeverReturnPartialBytes() {
        byte[] cabinet = ArchiveFixtures.StoredCab("file", new byte[] { 1, 2, 3 });
        cabinet[BitConverter.ToInt32(cabinet, 36)] = 1;
        using var source = new MemoryStream(cabinet);
        using ArchiveInspection inspection = ArchiveInspector.Open(source);
        Assert.True(Assert.Single(inspection.Entries).CanRead);
        Assert.Equal(1, inspection.Entries[0].FolderChecksumFrameCount);
        Assert.StartsWith("CAB_CHECKSUM:", Assert.Throws<InvalidDataException>(() => inspection.OpenRead(0)).Message);
        Assert.Equal(3, inspection.ReadBytes);
    }

    [Fact]
    public void SnapshotOwnsItsBytesAndReturnedMemberStreamsSurviveDisposal() {
        var options = new ArchiveInspectionOptions();
        using var openCancellation = new CancellationTokenSource();
        byte[] msi = ArchiveFixtures.Msi();
        using var source = new MemoryStream(msi, writable: true);
        source.Position = 9;
        ArchiveInspection inspection = ArchiveInspector.Open(source, options, openCancellation.Token);
        Assert.Equal(9, source.Position); Assert.True(source.CanRead);
        Array.Clear(msi, 0, msi.Length); source.Dispose(); openCancellation.Cancel();
        options.MaxEntryBytes = 1; options.MaxExpandedBytes = 1;
        using Stream member = inspection.OpenRead(0);
        ArchiveStorageStream raw = inspection.StorageStreams[0];
        using Stream storage = inspection.OpenStorageStream(raw.Ordinal);
        inspection.Dispose(); inspection.Dispose();
        Assert.Equal("hello", Text(member)); Assert.Equal(raw.Length, Bytes(storage).LongLength);
        Assert.Throws<ObjectDisposedException>(() => inspection.OpenRead(0));
        Assert.Throws<ObjectDisposedException>(() => inspection.OpenStorageStream(0));
    }

    private static string Text(Stream stream) { using var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: true); return reader.ReadToEnd(); }
    internal static byte[] Bytes(Stream stream) { using var result = new MemoryStream(); stream.CopyTo(result); return result.ToArray(); }
}
