using OfficeIMO.Archives.Compression;

namespace OfficeIMO.Archives.Tests;

public sealed class ZstandardArchiveTests {
    [Theory]
    [InlineData("text-level1.zst", "text.txt")]
    [InlineData("text-level3-check.zst", "text.txt")]
    [InlineData("text-level19-nocheck.zst", "text.txt")]
    [InlineData("mixed-level22.zst", "mixed.bin")]
    [InlineData("mixed-level-fast.zst", "mixed.bin")]
    [InlineData("mixed-window1k.zst", "mixed.bin")]
    [InlineData("mixed-blocks.zst", "mixed.bin")]
    [InlineData("empty.zst", "empty.bin")]
    public void DecodesFramesWrittenByTheReferenceTool(string fixture, string input) {
        byte[] expected = ReadInput(input);
        var budget = new ArchiveOperationBudget();
        var output = new MemoryStream();

        ZstandardReadResult result = ZstandardArchive.Decompress(Open(fixture), output, new ArchiveReadOptions { OperationBudget = budget });

        Assert.Equal(expected, output.ToArray());
        Assert.Equal(1, result.FrameCount);
        Assert.Equal(expected.LongLength, result.Length);
        Assert.Equal(expected.LongLength, budget.Usage.SelectedBytes);
    }

    [Fact]
    public void ConcatenatedFramesFormOneStreamAndSkippableFramesAreIgnored() {
        byte[] expected = ReadInput("text.txt").Concat(ReadInput("mixed.bin")).ToArray();
        var output = new MemoryStream();

        ZstandardReadResult result = ZstandardArchive.Decompress(Open("concatenated.zst"), output);

        Assert.Equal(expected, output.ToArray());
        Assert.Equal(2, result.FrameCount);
        Assert.Equal(1, result.SkippableFrameCount);
    }

    [Fact]
    public void NonseekableInputIsSnapshotted() {
        var output = new MemoryStream();
        ZstandardArchive.Decompress(new ForwardOnlyStream(ReadFixture("text-level3-check.zst")), output);
        Assert.Equal(ReadInput("text.txt"), output.ToArray());
    }

    [Fact]
    public void DeclaredWindowsAboveTheLimitAreRefusedBeforeDecoding() {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(Open("text-window128m.zst"), new MemoryStream()));
        Assert.Equal(nameof(ArchiveReadOptions.MaxWindowBytes), error.LimitName);

        var output = new MemoryStream();
        ZstandardArchive.Decompress(Open("text-window128m.zst"), output, new ArchiveReadOptions { MaxWindowBytes = 128L * 1024 * 1024 });
        Assert.Equal(ReadInput("text.txt"), output.ToArray());
    }

    [Fact]
    public void OutputBeyondTheEntryLimitIsRefused() {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(Open("mixed-blocks.zst"), new MemoryStream(), new ArchiveReadOptions { MaxEntryBytes = 64 * 1024 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxEntryBytes), error.LimitName);
    }

    [Fact]
    public void TrainedDictionaryFramesNeedTheMatchingDictionary() {
        ZstandardDictionary dictionary = ZstandardDictionary.Load(ReadFixture("trained.dict"));
        Assert.Equal(4242u, dictionary.Id);
        Assert.True(dictionary.HasEntropyTables);

        var output = new MemoryStream();
        ZstandardArchive.DecompressWithDictionary(Open("message-trained-dict.zst"), output, dictionary);
        Assert.Equal(ReadInput("message.json"), output.ToArray());

        ArchiveCorruptionException missing = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(Open("message-trained-dict.zst"), new MemoryStream()));
        Assert.Contains("4242", missing.Expected);
    }

    [Fact]
    public void RawContentDictionariesAreReferencedAsHistory() {
        ZstandardDictionary dictionary = ZstandardDictionary.Load(ReadFixture("raw-content.dict"));
        Assert.Equal(0u, dictionary.Id);
        Assert.False(dictionary.HasEntropyTables);

        var output = new MemoryStream();
        ZstandardArchive.DecompressWithDictionary(Open("message-raw-dict.zst"), output, dictionary);
        Assert.Equal(ReadInput("message.json"), output.ToArray());
    }

    [Fact]
    public void RawDictionariesDoNotSatisfyFramesThatNameADictionary() {
        ZstandardDictionary raw = ZstandardDictionary.Load(ReadFixture("raw-content.dict"));
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.DecompressWithDictionary(Open("message-trained-dict.zst"), new MemoryStream(), raw));
        Assert.Contains("4242", error.Expected);
    }

    [Fact]
    public void ADictionaryMagicWithoutAnIdIsRejected() {
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardDictionary.Load(new byte[] { 0x37, 0xA4, 0x30, 0xEC, 0x01 }));
        Assert.Equal(4, error.Offset);
    }

    [Fact]
    public void ATwoByteZeroSequenceCountEndsTheBlockAfterItsLiterals() {
        // A frame with a 1 KiB window and one compressed block holding raw literals "abc" and the count 0x80 0x00.
        byte[] source = { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x00, 0x35, 0x00, 0x00, 0x18, 0x61, 0x62, 0x63, 0x80, 0x00 };
        var output = new MemoryStream();
        ZstandardArchive.Decompress(new MemoryStream(source), output);
        Assert.Equal(Encoding.ASCII.GetBytes("abc"), output.ToArray());
    }

    [Fact]
    public void CopyingTheDictionaryIntoEachFrameCountsAsDecodeWork() {
        ZstandardDictionary dictionary = ZstandardDictionary.Load(Enumerable.Range(0, 1024 * 1024).Select(index => (byte)index).ToArray());
        // Twenty empty single-segment frames; each starts from a fresh copy of the 1 MiB dictionary.
        byte[] frame = { 0x28, 0xB5, 0x2F, 0xFD, 0x20, 0x00, 0x01, 0x00, 0x00 };
        byte[] source = Enumerable.Repeat(frame, 20).SelectMany(bytes => bytes).ToArray();
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.DecompressWithDictionary(new MemoryStream(source), new MemoryStream(), dictionary,
                new ArchiveReadOptions { MaxDecompressionWorkBytes = 8L * 1024 * 1024 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxDecompressionWorkBytes), error.LimitName);
    }

    [Fact]
    public void EveryTruncationFailsWithATypedErrorInsideTheInput() {
        byte[] source = ReadFixture("text-level3-check.zst");
        for (int length = 0; length < source.Length; length += length < 1024 ? 1 : 7) {
            ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
                ZstandardArchive.Decompress(new MemoryStream(source, 0, length, writable: false), new MemoryStream()));
            Assert.Equal("Zstandard", error.Format);
            Assert.InRange(error.Offset, 0, length);
        }
    }

    [Theory]
    [InlineData("text-level3-check.zst", "text.txt", true)]
    [InlineData("text-level19-nocheck.zst", "text.txt", false)]
    [InlineData("mixed-window1k.zst", "mixed.bin", true)]
    public void FlippedBitsFailWithTypedErrorsOnly(string fixture, string input, bool checksummed) {
        byte[] original = ReadFixture(fixture);
        byte[] expected = ReadInput(input);
        var random = new Random(7);
        for (int trial = 0; trial < 1000; trial++) {
            byte[] damaged = (byte[])original.Clone();
            int position = trial < 64 ? trial : random.Next(damaged.Length);
            damaged[position] ^= (byte)(1 << random.Next(8));
            var output = new MemoryStream();
            Exception? error = Record.Exception(() =>
                ZstandardArchive.Decompress(new MemoryStream(damaged, writable: false), output,
                    new ArchiveReadOptions { MaxEntryBytes = 1024 * 1024, MaxWindowBytes = 1024 * 1024 }));
            // Header bits the format leaves unused, or a larger declared window, may decode unchanged; the content
            // checksum catches every other change, and the failure must be a typed corruption or limit error.
            if (error == null) {
                // Without a content checksum a damaged literal can decode to different bytes of the right length.
                if (checksummed) Assert.Equal(expected, output.ToArray());
                continue;
            }
            Assert.True(error is ArchiveCorruptionException or OfficeInspectionResourceLimitException,
                $"Flipping a bit at byte {position} raised {error.GetType().Name}: {error.Message}");
            if (error is ArchiveCorruptionException corruption) Assert.InRange(corruption.Offset, 0, damaged.Length);
        }
    }

    [Fact]
    public void TrailingBytesAfterTheLastFrameAreReportedAtTheirOffset() {
        byte[] frame = ReadFixture("text-level1.zst");
        byte[] source = frame.Concat(new byte[] { 1, 2, 3, 4 }).ToArray();
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(frame.Length, error.Offset);
    }

    [Fact]
    public void ReservedBlockTypeIsRejectedAtTheBlockHeader() {
        // Frame header: magic, descriptor with single segment and a 1-byte content size of 1, then a last block of type 3.
        byte[] source = { 0x28, 0xB5, 0x2F, 0xFD, 0x20, 0x01, 0x07, 0x00, 0x00 };
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(6, error.Offset);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(31)]
    [InlineData(32)]
    [InlineData(33)]
    [InlineData(1000)]
    public void XxHash64MatchesOneShotAcrossChunkBoundaries(int length) {
        byte[] data = Enumerable.Range(0, length).Select(index => (byte)(index * 31 + 7)).ToArray();
        ulong whole = XxHash64.Compute(data, 0, data.Length);
        var hash = new XxHash64();
        for (int offset = 0; offset < data.Length; offset += 13) hash.Append(data, offset, Math.Min(13, data.Length - offset));
        Assert.Equal(whole, hash.Complete());
        if (length == 0) Assert.Equal(0xEF46DB3751D8E999UL, whole);
    }

    private static Stream Open(string fixture) => new MemoryStream(ReadFixture(fixture), writable: false);

    private static byte[] ReadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentZstandard", name));

    private static byte[] ReadInput(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentZstandard", "inputs", name));

    private sealed class ForwardOnlyStream : Stream {
        private readonly MemoryStream _inner;
        public ForwardOnlyStream(byte[] data) => _inner = new MemoryStream(data, writable: false);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
