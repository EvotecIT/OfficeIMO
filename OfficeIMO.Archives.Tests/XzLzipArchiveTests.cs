using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives.Tests;

public sealed class XzLzipArchiveTests {
    [Theory]
    [InlineData("text-preset0.xz", "text.txt", XzCheck.Crc64, 1)]
    [InlineData("text-preset6-crc64.xz", "text.txt", XzCheck.Crc64, 1)]
    [InlineData("text-crc32.xz", "text.txt", XzCheck.Crc32, 1)]
    [InlineData("text-sha256.xz", "text.txt", XzCheck.Sha256, 1)]
    [InlineData("text-nocheck.xz", "text.txt", XzCheck.None, 1)]
    [InlineData("mixed-preset9e.xz", "mixed.bin", XzCheck.Crc64, 1)]
    [InlineData("mixed-blocks.xz", "mixed.bin", XzCheck.Crc64, 6)]
    [InlineData("mixed-threads.xz", "mixed.bin", XzCheck.Crc64, 3)]
    [InlineData("samples-delta.xz", "samples.bin", XzCheck.Crc64, 1)]
    [InlineData("text-dict4k.xz", "text.txt", XzCheck.Crc64, 1)]
    [InlineData("empty.xz", "empty.bin", XzCheck.Crc64, 0)]
    public void DecodesXzFilesWrittenByXzUtils(string fixture, string input, XzCheck check, int blocks) {
        byte[] expected = ReadInput(input);
        var budget = new ArchiveOperationBudget();
        var output = new MemoryStream();

        XzReadResult result = XzArchive.Decompress(Open(fixture), output, new ArchiveReadOptions { OperationBudget = budget });

        Assert.Equal(expected, output.ToArray());
        Assert.Equal(check, result.Check);
        Assert.Equal(blocks, result.BlockCount);
        Assert.Equal(1, result.StreamCount);
        Assert.Equal(expected.LongLength, budget.Usage.SelectedBytes);
    }

    [Fact]
    public void ConcatenatedXzStreamsWithPaddingFormOneStream() {
        var output = new MemoryStream();
        XzReadResult result = XzArchive.Decompress(Open("concatenated.xz"), output);
        Assert.Equal(ReadInput("text.txt").Concat(ReadInput("mixed.bin")).ToArray(), output.ToArray());
        Assert.Equal(2, result.StreamCount);
    }

    [Fact]
    public void DictionariesAboveTheWindowLimitAreRefusedBeforeDecoding() {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            XzArchive.Decompress(Open("text-dict128m.xz"), new MemoryStream()));
        Assert.Equal(nameof(ArchiveReadOptions.MaxWindowBytes), error.LimitName);

        var output = new MemoryStream();
        XzArchive.Decompress(Open("text-dict128m.xz"), output, new ArchiveReadOptions { MaxWindowBytes = 128L * 1024 * 1024 });
        Assert.Equal(ReadInput("text.txt"), output.ToArray());
    }

    [Fact]
    public void BranchFiltersAreReportedAsUnsupported() {
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => XzArchive.Decompress(Open("text-x86.xz"), new MemoryStream()));
        Assert.Contains("0x04", error.Message);
    }

    [Theory]
    [InlineData("text-level0.lz", "text.txt")]
    [InlineData("text-level6.lz", "text.txt")]
    [InlineData("mixed-level9.lz", "mixed.bin")]
    [InlineData("text-dict4k.lz", "text.txt")]
    [InlineData("empty.lz", "empty.bin")]
    public void DecodesLzipFilesWrittenByLzip(string fixture, string input) {
        byte[] expected = ReadInput(input);
        var output = new MemoryStream();
        LzipReadResult result = LzipArchive.Decompress(Open(fixture), output);
        Assert.Equal(expected, output.ToArray());
        Assert.Equal(1, result.MemberCount);
    }

    [Fact]
    public void ConcatenatedLzipMembersFormOneStream() {
        var output = new MemoryStream();
        LzipReadResult result = LzipArchive.Decompress(Open("concatenated.lz"), output);
        Assert.Equal(ReadInput("text.txt").Concat(ReadInput("mixed.bin")).ToArray(), output.ToArray());
        Assert.Equal(2, result.MemberCount);
    }

    [Theory]
    [InlineData("text-preset6-crc64.xz")]
    [InlineData("text-level6.lz")]
    public void OutputBeyondTheEntryLimitIsRefused(string fixture) {
        var options = new ArchiveReadOptions { MaxEntryBytes = 64 * 1024 };
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => Decompress(fixture, options));
        Assert.Equal(nameof(ArchiveReadOptions.MaxEntryBytes), error.LimitName);
    }

    [Theory]
    [InlineData("mixed-blocks.xz", "xz")]
    [InlineData("text-level6.lz", "lzip")]
    public void EveryTruncationFailsWithATypedErrorInsideTheInput(string fixture, string format) {
        byte[] source = ReadFixture(fixture);
        for (int length = 0; length < source.Length; length += length < 1024 ? 1 : 11) {
            ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
                Decompress(fixture, null, new MemoryStream(source, 0, length, writable: false)));
            Assert.Equal(format, error.Format);
            Assert.InRange(error.Offset, 0, length);
        }
    }

    [Theory]
    [InlineData("mixed-blocks.xz", "mixed.bin")]
    [InlineData("text-sha256.xz", "text.txt")]
    [InlineData("samples-delta.xz", "samples.bin")]
    [InlineData("text-level6.lz", "text.txt")]
    [InlineData("text-dict4k.lz", "text.txt")]
    [InlineData("text-dict4k.xz", "text.txt")]
    public void FlippedBitsFailWithTypedErrorsOnly(string fixture, string input) {
        byte[] original = ReadFixture(fixture);
        byte[] expected = ReadInput(input);
        var random = new Random(11);
        var options = new ArchiveReadOptions { MaxEntryBytes = 1024 * 1024, MaxWindowBytes = 16L * 1024 * 1024 };
        for (int trial = 0; trial < 400; trial++) {
            byte[] damaged = (byte[])original.Clone();
            int position = trial < 48 ? trial : random.Next(damaged.Length);
            damaged[position] ^= (byte)(1 << random.Next(8));
            var output = new MemoryStream();
            Exception? error = Record.Exception(() => Decompress(fixture, options, new MemoryStream(damaged, writable: false), output));
            // Every damaged byte is covered by a CRC or the block check; bits a decoder ignores must leave the output intact.
            if (error == null) {
                Assert.Equal(expected, output.ToArray());
                continue;
            }
            Assert.True(error is ArchiveCorruptionException or OfficeInspectionResourceLimitException,
                $"Flipping a bit at byte {position} raised {error.GetType().Name}: {error.Message}");
            if (error is ArchiveCorruptionException corruption) Assert.InRange(corruption.Offset, 0, damaged.Length);
        }
    }

    [Fact]
    public void TrailingDataAfterTheLastLzipMemberIsRejectedAtItsOffset() {
        byte[] member = ReadFixture("text-level0.lz");
        byte[] source = member.Concat(new byte[] { 0, 0, 0, 0 }).ToArray();
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            LzipArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(member.Length, error.Offset);
    }

    [Fact]
    public void XzStreamPaddingMustComeInFourByteUnits() {
        byte[] stream = ReadFixture("text-preset0.xz");
        byte[] source = stream.Concat(new byte[] { 0, 0 }).ToArray();
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            XzArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(stream.Length, error.Offset);
    }

    [Fact]
    public void Crc64MatchesTheXzCheckValue() {
        // The CHECK value from the xz format specification: CRC-64 of the nine ASCII digits "123456789".
        byte[] digits = Encoding.ASCII.GetBytes("123456789");
        Assert.Equal(0x995DC9BBDF1939FAUL, Crc64.Compute(digits, 0, digits.Length));
    }

    private static object Decompress(string fixture, ArchiveReadOptions? options, Stream? source = null, Stream? output = null) {
        source ??= Open(fixture);
        output ??= new MemoryStream();
        return fixture.EndsWith(".xz", StringComparison.Ordinal)
            ? XzArchive.Decompress(source, output, options)
            : LzipArchive.Decompress(source, output, options);
    }

    private static Stream Open(string fixture) => new MemoryStream(ReadFixture(fixture), writable: false);

    private static byte[] ReadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentXzLzip", name));

    private static byte[] ReadInput(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentXzLzip", "inputs", name));
}
