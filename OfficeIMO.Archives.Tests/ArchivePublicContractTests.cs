using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchivePublicContractTests {
    [Fact]
    public void CaseDistinctEmbeddedCabinetsKeepSeparateFileMappings() {
        using var fixture = new ArchiveFixtures();
        byte[] package = RewriteMsi(ArchiveFixtures.Msi(twoCabinets: true), streams => {
            ReplaceString(streams, "#second.cab", "#PAYLOAD.CAB");
            string previous = MsiPackageReader.EncodeStreamName("second.cab", false);
            byte[] cabinet = streams[previous]; streams.Remove(previous);
            streams.Add(MsiPackageReader.EncodeStreamName("PAYLOAD.CAB", false), cabinet);
        });
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination);
        Assert.Equal(new[] { "PolicyDefinitions/en-US/Second.adml", "PolicyDefinitions/en-US/Template.adml" }, result.Files.Select(file => file.RelativePath));
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
        Assert.Equal("world", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Second.adml")));
    }

    [Fact]
    public void WronglyCasedEmbeddedCabinetReferenceIsMissing() {
        using var fixture = new ArchiveFixtures();
        byte[] package = RewriteMsi(ArchiveFixtures.Msi(), streams => ReplaceString(streams, "#payload.cab", "#PAYLOAD.CAB"));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false, 2, "CAB_COMPRESSION")]
    [InlineData(true, 2, "CAB_COMPRESSION")]
    [InlineData(false, 0x0e03, "CAB_LZX_WINDOW")]
    [InlineData(true, 0x0e03, "CAB_LZX_WINDOW")]
    [InlineData(false, -1, "CAB_MULTIPART")]
    [InlineData(true, -1, "CAB_MULTIPART")]
    public void UnsupportedCabinetFeaturesRemainDistinguishable(bool embedded, int compression, string code) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = ArchiveFixtures.StoredCab("PayloadKey", Encoding.UTF8.GetBytes("hello"));
        if (compression < 0) cabinet = WithPreviousCabinet(cabinet);
        else { cabinet[42] = (byte)compression; cabinet[43] = (byte)(compression >> 8); }
        byte[] package = embedded ? ReplaceCabinet(cabinet) : cabinet;
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.StartsWith(code + ":", exception.Message);
        Assert.NotNull(exception.InnerException);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MalformedCabinetsRemainInvalidData(bool embedded, bool checksum) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = ArchiveFixtures.StoredCab("PayloadKey", Encoding.UTF8.GetBytes("hello"));
        if (checksum) cabinet[BitConverter.ToInt32(cabinet, 36)] = 1;
        else Array.Resize(ref cabinet, cabinet.Length - 1);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(embedded ? ReplaceCabinet(cabinet) : cabinet), fixture.Destination));
        Assert.StartsWith(checksum ? "CAB_CHECKSUM:" : "CAB_HEADER:", exception.Message);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void EmbeddedCabinetOpensReuseTheCompoundDirectoryAndAllocationIndex() {
        using var source = new CountingStream(ArchiveFixtures.Msi(cabinetCount: 16, emptyPayloads: true));
        using var database = new MsiPackageReader(source, new ArchiveLimits(new ArchiveExtractionOptions()), CancellationToken.None);
        IReadOnlyList<MsiPayloadFile> files = database.ReadFiles();
        long before = source.BytesRead;
        foreach (MsiPayloadFile file in files) {
            using Stream cabinet = database.OpenCabinet(file.Cabinet);
            Assert.True(cabinet.Length > 0);
        }
        Assert.Equal(0, source.BytesRead - before);
    }

    [Theory]
    [InlineData(16, 16, true)]
    [InlineData(16, 15, false)]
    [InlineData(257, 256, false)]
    public void ZeroByteCabinetsObserveCountLimitsWhenEveryFileIsUnselected(int count, int limit, bool accepted) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.Msi(cabinetCount: count, emptyPayloads: true));
        var options = new ArchiveExtractionOptions { MaxCabinetCount = limit, IncludeFile = _ => false };
        if (accepted) {
            ArchiveExtractionResult result = ArchiveExtractor.Extract(package, fixture.Destination, options);
            Assert.Empty(result.Files); Assert.Equal(0, result.TotalBytes);
        } else {
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination, options));
            Assert.Contains("embedded-cabinet count", exception.Message);
            Assert.False(Directory.Exists(fixture.Destination));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CompoundSessionDisposalInvalidatesViewsAndRespectsSourceOwnership(bool leaveOpen) {
        var source = new CountingStream(ArchiveFixtures.Msi());
        var session = new OfficeCompoundFileReader.ReadSession(source, new OfficeCompoundReadOptions(), leaveOpen, CancellationToken.None);
        using Stream cabinet = session.OpenStream(MsiPackageReader.EncodeStreamName("payload.cab", false));
        Assert.Equal((int)'M', cabinet.ReadByte());
        session.Dispose();
        Assert.Throws<ObjectDisposedException>(() => cabinet.ReadByte());
        Assert.Throws<ObjectDisposedException>(() => session.OpenStream(MsiPackageReader.EncodeStreamName("payload.cab", false)));
        Assert.Equal(leaveOpen, source.CanRead);
        source.Dispose();
    }

    private static byte[] ReplaceCabinet(byte[] cabinet) => RewriteMsi(ArchiveFixtures.Msi(), streams =>
        streams[MsiPackageReader.EncodeStreamName("payload.cab", false)] = cabinet);

    private static byte[] RewriteMsi(byte[] original, Action<Dictionary<string, byte[]>> rewrite) {
        Assert.True(OfficeCompoundFileReader.TryRead(original, out Dictionary<string, byte[]> streams, out string? error), error);
        rewrite(streams);
        return OfficeCompoundFileWriter.Write(streams.Select(stream => new OfficeCompoundStream(stream.Key, stream.Value)).ToArray(),
            new Guid("000c1084-0000-0000-c000-000000000046"));
    }

    private static void ReplaceString(Dictionary<string, byte[]> streams, string previous, string replacement) {
        string poolName = MsiPackageReader.EncodeStreamName("_StringPool", true);
        string dataName = MsiPackageReader.EncodeStreamName("_StringData", true);
        byte[] pool = streams[poolName]; byte[] data = streams[dataName];
        using var output = new MemoryStream(); int offset = 0;
        for (int record = 4; record < pool.Length; record += 4) {
            int length = pool[record] | pool[record + 1] << 8;
            string value = Encoding.UTF8.GetString(data, offset, length);
            byte[] bytes = Encoding.UTF8.GetBytes(value == previous ? replacement : value);
            pool[record] = (byte)bytes.Length; pool[record + 1] = (byte)(bytes.Length >> 8);
            output.Write(bytes, 0, bytes.Length); offset += length;
        }
        streams[dataName] = output.ToArray();
    }

    private static byte[] WithPreviousCabinet(byte[] cabinet) {
        byte[] names = Encoding.ASCII.GetBytes("previous.cab\0disk\0");
        byte[] result = new byte[cabinet.Length + names.Length];
        Buffer.BlockCopy(cabinet, 0, result, 0, 36);
        Buffer.BlockCopy(names, 0, result, 36, names.Length);
        Buffer.BlockCopy(cabinet, 36, result, 36 + names.Length, cabinet.Length - 36);
        Buffer.BlockCopy(BitConverter.GetBytes(result.Length), 0, result, 8, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(BitConverter.ToInt32(cabinet, 16) + names.Length), 0, result, 16, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(BitConverter.ToInt32(cabinet, 36) + names.Length), 0, result, 36 + names.Length, 4);
        result[30] = 1; result[34] = 1;
        return result;
    }

    private sealed class CountingStream : MemoryStream {
        internal CountingStream(byte[] bytes) : base(bytes, writable: false) { }
        internal long BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) { int read = base.Read(buffer, offset, count); BytesRead += read; return read; }
    }
}
