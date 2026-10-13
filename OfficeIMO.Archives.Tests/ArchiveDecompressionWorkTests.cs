namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveDecompressionWorkTests {
    private const long FrameFloor = 32768;

    [Fact]
    public void TinyMsZipFramesConsumeTheMinimumWorkEvenWithLittleOutput() {
        byte[] cabinet = MsZipCab(new[] { new byte[] { 1 }, new byte[] { 2 }, new byte[] { 3 }, new byte[] { 4 } });
        using var fixture = new ArchiveFixtures();

        AssertWorkLimit(fixture, cabinet, 3 * FrameFloor, maxExpandedBytes: 4);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OverlappingFoldersAndUnselectedGapsShareTheWorkBudget(bool selected) {
        // Only one byte in folder zero is a member. Both two-byte folders must still
        // be decoded, even when they refer to the same physical CFDATA records.
        byte[] cabinet = MsZipCab(new[] { new byte[] { 1 }, new byte[] { 2 } }, folderCount: 2, memberLength: 1);
        using var fixture = new ArchiveFixtures();

        AssertWorkLimit(fixture, cabinet, 3 * FrameFloor, maxExpandedBytes: 4, selected);
    }

    [Fact]
    public void MsZipDictionaryReplayCountsBeyondTheMinimumFrameCharge() {
        byte[] first = Enumerable.Range(0, 32768).Select(index => (byte)(index % 251)).ToArray();
        byte[] cabinet = MsZipCab(new[] { first, new byte[] { 42 } });
        // First record: 8 CFDATA + 7 MSZIP/stored-DEFLATE + 32,768 payload.
        // Second frame: 32,768 dictionary replay + one newly decoded byte.
        const long exactWork = 32783 + 32769;
        using (var rejected = new ArchiveFixtures()) {
            AssertWorkLimit(rejected, cabinet, exactWork - 1, maxExpandedBytes: 32769);
        }
        using var accepted = new ArchiveFixtures();

        ArchiveExtractionResult result = ArchiveExtractor.Extract(accepted.Package(cabinet), accepted.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 32769, MaxDecompressionWorkBytes = exactWork });

        Assert.Equal(32769, result.TotalBytes);
        Assert.Equal(first.Concat(new byte[] { 42 }), File.ReadAllBytes(Path.Combine(accepted.Destination, "template.admx")));
    }

    [Fact]
    public void CompressedInputAndReserveBytesCountBeforeChecksumValidation() {
        byte[] cabinet = MsZipCab(new[] { new byte[32768] }, dataReserve: 3);
        // The payload alone expands to exactly 32 KiB; its record plus reserve is 32,786 bytes.
        using var fixture = new ArchiveFixtures();
        AssertWorkLimit(fixture, cabinet, 32785, maxExpandedBytes: 32768);

        cabinet[cabinet.Length - 1] ^= 1;
        using var corrupted = new ArchiveFixtures();
        AssertWorkLimit(corrupted, cabinet, 32785, maxExpandedBytes: 32768);
        using var admitted = new ArchiveFixtures();
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            admitted.Package(cabinet), admitted.Destination,
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = 32786 }));
        Assert.Contains("CAB_CHECKSUM", exception.Message);
    }

    [Fact]
    public void LzxFramesUseTheSamePreflightBudgetAndExactBoundary() {
        // One LZX uncompressed block: no Intel transform, type 3, length 1,
        // word padding, repeated offsets R0/R1/R2 = 1, literal 'A', byte padding.
        byte[] chunk = { 0, 0x30, 0x10, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, (byte)'A', 0 };
        byte[] cabinet = ArchiveFixtures.StoredCab("template.admx", chunk);
        cabinet[42] = 3; cabinet[43] = 15; // LZX, 32 KiB window.
        cabinet[44] = 1; // The member expands to one byte.
        int dataOffset = BitConverter.ToInt32(cabinet, 36);
        cabinet[dataOffset + 6] = 1;
        using (var rejected = new ArchiveFixtures()) {
            AssertWorkLimit(rejected, cabinet, FrameFloor - 1, maxExpandedBytes: 1);
        }
        using var accepted = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(accepted.Package(cabinet), accepted.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 1, MaxDecompressionWorkBytes = FrameFloor });
        Assert.Equal(1, result.TotalBytes);
        Assert.Equal(new byte[] { (byte)'A' }, File.ReadAllBytes(Path.Combine(accepted.Destination, "template.admx")));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmbeddedMsiCabinetsShareTheWorkBudgetIncludingUnselectedFiles(bool selected) {
        using var fixture = new ArchiveFixtures();
        AssertWorkLimit(fixture, ArchiveFixtures.Msi(twoCabinets: true), FrameFloor, maxExpandedBytes: 10, selected);
    }

    [Fact]
    public void ExactMsiWorkBudgetAcceptsBothCabinets() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)),
            fixture.Destination, new ArchiveExtractionOptions { MaxExpandedBytes = 10, MaxDecompressionWorkBytes = 2 * FrameFloor });
        Assert.Equal(10, result.TotalBytes);
        Assert.Equal(2, result.Files.Count);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(0x0f03, true)]
    [InlineData(1, false)]
    public void ExactWorkAndExpansionBudgetsAllowFollowingZeroFrameCabinet(int emptyCompression, bool selected) {
        using var fixture = new ArchiveFixtures();
        byte[] package = ArchiveFixtures.Msi(twoCabinets: true, emptySecondPayload: true, emptySecondCompression: emptyCompression);
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 5, MaxDecompressionWorkBytes = FrameFloor, IncludeFile = _ => selected });
        Assert.Equal(selected ? 5 : 0, result.TotalBytes);
        Assert.Equal(selected ? 2 : 0, result.Files.Count);
    }

    [Theory]
    [InlineData(0L)]
    [InlineData(-1L)]
    public void PublicWorkBudgetMustBePositive(long budget) {
        using var fixture = new ArchiveFixtures();
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.EmptyStoredCab("empty")), fixture.Destination,
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = budget }));
        Assert.Equal(nameof(ArchiveExtractionOptions.MaxDecompressionWorkBytes), exception.ParamName);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void LargeWorkBudgetAcceptsNonemptyCabinetsWithoutOverflow() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)), fixture.Destination,
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = long.MaxValue });
        Assert.Equal(10, result.TotalBytes);
        Assert.Equal(2, result.Files.Count);
    }

    [Fact]
    public void AZeroFrameCabinetConsumesNoWork() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.EmptyStoredCab("empty")), fixture.Destination,
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = 1 });
        Assert.Equal(0, result.TotalBytes);
        Assert.Equal("empty", Assert.Single(result.Files).RelativePath);
        Assert.Empty(File.ReadAllBytes(Path.Combine(fixture.Destination, "empty")));
    }

    private static void AssertWorkLimit(ArchiveFixtures fixture, byte[] package, long work, long maxExpandedBytes, bool selected = true) {
        OfficeInspectionResourceLimitException exception = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveExtractor.Extract(
            fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = maxExpandedBytes, MaxDecompressionWorkBytes = work, IncludeFile = _ => selected }));
        Assert.Equal("MaxDecompressionWorkBytes", exception.LimitName);
        Assert.Equal(work, exception.Maximum);
        Assert.Equal("CAB_WORK_LIMIT", Assert.IsType<OfficeIMO.Core.Internal.Cabinet.OfficeCabinetException>(exception.InnerException).Code);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.DoesNotContain(Directory.EnumerateDirectories(fixture.Root), path => Path.GetFileName(path).StartsWith(".officeimo-archive-", StringComparison.Ordinal));
    }

    private static byte[] MsZipCab(IReadOnlyList<byte[]> payloads, int folderCount = 1, int? memberLength = null, byte dataReserve = 0) {
        byte[] name = Encoding.ASCII.GetBytes("template.admx\0");
        int foldersOffset = dataReserve == 0 ? 36 : 40;
        int filesOffset = foldersOffset + 8 * folderCount;
        int dataOffset = filesOffset + 16 + name.Length;
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(new byte[dataOffset]);
        foreach (byte[] payload in payloads) {
            using var record = new MemoryStream();
            using var recordWriter = new BinaryWriter(record, Encoding.UTF8, leaveOpen: true);
            recordWriter.Write((ushort)(payload.Length + 7));
            recordWriter.Write((ushort)payload.Length);
            recordWriter.Write(new byte[dataReserve]);
            recordWriter.Write((byte)'C'); recordWriter.Write((byte)'K');
            recordWriter.Write((byte)1); // Final RFC 1951 stored block, no history references.
            recordWriter.Write((ushort)payload.Length);
            recordWriter.Write(unchecked((ushort)~payload.Length));
            recordWriter.Write(payload);
            byte[] bytes = record.ToArray();
            uint checksum = Checksum(bytes, 4 + dataReserve, bytes.Length - 4 - dataReserve);
            checksum = Checksum(bytes, 0, 4 + dataReserve, checksum);
            writer.Write(checksum);
            writer.Write(bytes);
        }
        stream.Position = 0;
        writer.Write(Encoding.ASCII.GetBytes("MSCF")); writer.Write(0u); writer.Write((uint)stream.Length);
        writer.Write(0u); writer.Write((uint)filesOffset); writer.Write(0u);
        writer.Write((byte)3); writer.Write((byte)1); writer.Write((ushort)folderCount); writer.Write((ushort)1);
        writer.Write((ushort)(dataReserve == 0 ? 0 : 4)); writer.Write(0u);
        if (dataReserve != 0) { writer.Write((ushort)0); writer.Write((byte)0); writer.Write(dataReserve); }
        for (int folder = 0; folder < folderCount; folder++) {
            writer.Write((uint)dataOffset); writer.Write((ushort)payloads.Count); writer.Write((ushort)1);
        }
        writer.Write((uint)(memberLength ?? payloads.Sum(payload => payload.Length)));
        writer.Write(0u); writer.Write((ushort)0); writer.Write(0u); writer.Write((ushort)0); writer.Write(name);
        return stream.ToArray();
    }

    private static uint Checksum(byte[] bytes, int offset, int count, uint seed = 0) {
        int end = offset + count;
        while (offset <= end - 4) {
            seed ^= (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
            offset += 4;
        }
        uint tail = 0;
        while (offset < end) tail = (tail << 8) | bytes[offset++];
        return seed ^ tail;
    }
}
