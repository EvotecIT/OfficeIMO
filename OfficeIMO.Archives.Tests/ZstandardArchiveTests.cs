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
    // Through a 1 KiB window the output outgrows the first history buffer, so later matches read slid history.
    [InlineData("text4-window1k.zst", "text.txt x4")]
    public void DecodesFramesWrittenByTheReferenceTool(string fixture, string input) {
        byte[] expected = Expected(input);
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

    public static TheoryData<string, string, string?, bool> Corpus() => new() {
        { "text-level1.zst", "text.txt", null, true },
        { "text-level3-check.zst", "text.txt", null, true },
        { "text-level19-nocheck.zst", "text.txt", null, false },
        { "mixed-level22.zst", "mixed.bin", null, true },
        { "mixed-level-fast.zst", "mixed.bin", null, true },
        { "mixed-window1k.zst", "mixed.bin", null, true },
        { "mixed-blocks.zst", "mixed.bin", null, true },
        { "empty.zst", "empty.bin", null, true },
        { "concatenated.zst", "concatenated", null, true },
        { "text-window128m.zst", "text.txt", null, true },
        { "message-trained-dict.zst", "message.json", "trained.dict", true },
        { "message-raw-dict.zst", "message.json", "raw-content.dict", true }
    };

    [Theory]
    [MemberData(nameof(Corpus))]
    public void EveryTruncationFailsWithATypedErrorInsideTheInput(string fixture, string input, string? dictionaryName, bool checksummed) {
        _ = checksummed;
        byte[] source = ReadFixture(fixture);
        byte[] expected = Expected(input);
        ZstandardDictionary? dictionary = dictionaryName == null ? null : ZstandardDictionary.Load(ReadFixture(dictionaryName));
        for (int length = 0; length < source.Length; length++) {
            var output = new MemoryStream();
            Exception? error = Record.Exception(() => Decode(new MemoryStream(source, 0, length, writable: false), output, dictionary));
            if (error == null) {
                // Only a cut between complete frames of a concatenated stream is itself a valid stream.
                Assert.Equal("concatenated.zst", fixture);
                Assert.Equal(expected.Take((int)output.Length), output.ToArray());
                continue;
            }
            ArchiveCorruptionException corruption = Assert.IsType<ArchiveCorruptionException>(error);
            Assert.Equal("Zstandard", corruption.Format);
            Assert.InRange(corruption.Offset, 0, length);
        }
    }

    [Theory]
    [MemberData(nameof(Corpus))]
    public void FlippedBitsFailWithTypedErrorsOnly(string fixture, string input, string? dictionaryName, bool checksummed) {
        byte[] original = ReadFixture(fixture);
        byte[] expected = Expected(input);
        ZstandardDictionary? dictionary = dictionaryName == null ? null : ZstandardDictionary.Load(ReadFixture(dictionaryName));
        var random = new Random(7);
        int trials = Math.Min(1000, 8 * original.Length);
        for (int trial = 0; trial < trials; trial++) {
            byte[] damaged = (byte[])original.Clone();
            int position = trial < Math.Min(64, damaged.Length) ? trial : random.Next(damaged.Length);
            damaged[position] ^= (byte)(1 << random.Next(8));
            var output = new MemoryStream();
            Exception? error = Record.Exception(() => Decode(new MemoryStream(damaged, writable: false), output, dictionary,
                new ArchiveReadOptions { MaxEntryBytes = 1024 * 1024, MaxWindowBytes = 128L * 1024 * 1024 }));
            // Header bits the format leaves unused, a skippable frame's payload, or a larger declared window may decode
            // unchanged; the content checksum catches every other change, and the failure must be a typed corruption or
            // limit error.
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

    private static ZstandardReadResult Decode(Stream source, Stream output, ZstandardDictionary? dictionary, ArchiveReadOptions? options = null) {
        options ??= new ArchiveReadOptions { MaxWindowBytes = 128L * 1024 * 1024 };
        return dictionary == null
            ? ZstandardArchive.Decompress(source, output, options)
            : ZstandardArchive.DecompressWithDictionary(source, output, dictionary, options);
    }

    private static byte[] Expected(string input) => input switch {
        "concatenated" => ReadInput("text.txt").Concat(ReadInput("mixed.bin")).ToArray(),
        "text.txt x4" => Enumerable.Repeat(ReadInput("text.txt"), 4).SelectMany(bytes => bytes).ToArray(),
        _ => ReadInput(input)
    };

    [Fact]
    public void TheSourceCannotAlsoBeTheDestination() {
        var stream = new MemoryStream();
        stream.Write(ReadFixture("text-level1.zst"), 0, ReadFixture("text-level1.zst").Length);
        stream.Position = 0;
        Assert.Throws<ArgumentException>(() => ZstandardArchive.Decompress(stream, stream));
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
    public void ATrailingSkippableFrameLeavesTheSourceAfterIt() {
        byte[] frame = ReadFixture("text-level1.zst");
        byte[] skippable = { 0x50, 0x2A, 0x4D, 0x18, 0x05, 0x00, 0x00, 0x00, 1, 2, 3, 4, 5 };
        var source = new MemoryStream(frame.Concat(skippable).ToArray());

        ZstandardReadResult result = ZstandardArchive.Decompress(source, new MemoryStream());

        Assert.Equal(1, result.SkippableFrameCount);
        Assert.Equal(source.Length, source.Position);
    }

    [Fact]
    public void ASourcePositionedPastItsEndIsEmptyInput() {
        var source = new MemoryStream(ReadFixture("text-level1.zst"));
        source.Position = source.Length + 1;
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(source, new MemoryStream()));
        Assert.Equal(0, error.Offset);
    }

    [Fact]
    public void HistoryThatStopsGrowingShortOfABlockSlidesInstead() {
        // A 704 KiB window (exponent 9, mantissa 3) grows its history to a 1,056 KiB cap. After eight full raw blocks the
        // capped resize leaves only 32 KiB free, so the ninth block needs a slide.
        const int block = 128 * 1024;
        var expected = new byte[12 * block];
        for (int index = 0; index < expected.Length; index++) expected[index] = (byte)(index * 7 + index / block);
        var frame = new MemoryStream();
        frame.Write(new byte[] { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x4B }, 0, 6);
        for (int index = 0; index < 12; index++) {
            int header = (block << 3) | (index == 11 ? 1 : 0);
            frame.Write(new[] { (byte)header, (byte)(header >> 8), (byte)(header >> 16) }, 0, 3);
            frame.Write(expected, index * block, block);
        }
        frame.Position = 0;
        var output = new MemoryStream();

        ZstandardArchive.Decompress(frame, output);

        Assert.Equal(expected, output.ToArray());
    }

    [Fact]
    public void LoadRawKeepsContentThatStartsWithTheDictionaryMagic() {
        byte[] content = { 0x37, 0xA4, 0x30, 0xEC, 0x01, 0x00, 0x00, 0x00, 0x2A, 0x2A };
        ZstandardDictionary dictionary = ZstandardDictionary.LoadRaw(content);
        Assert.Equal(0u, dictionary.Id);
        Assert.False(dictionary.HasEntropyTables);
        Assert.Equal(content.Length, dictionary.ContentLength);
    }

    [Fact]
    public void FourHuffmanStreamsNeedAtLeastSixLiterals() {
        // A 1 KiB window and one compressed block whose literals header declares four streams but only four literals.
        byte[] source = { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x00, 0x25, 0x00, 0x00, 0x46, 0x40, 0x00, 0x00 };
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(9, error.Offset);
    }

    [Theory]
    [InlineData(10)]
    [InlineData(9 + 1024)]
    public void CancellationStopsSeekableReadsInsideABlock(long cancelAt) {
        // A 1 KiB window and one raw 1 KiB block, read one byte at a time. Cancellation arrives with the first payload byte
        // or with the read that completes the block; either way nothing is written.
        var frame = new List<byte> { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x00, 0x01, 0x20, 0x00 };
        frame.AddRange(Enumerable.Repeat((byte)0x41, 1024));
        using var cancellation = new CancellationTokenSource();
        var source = new TrickleStream(frame.ToArray(), cancelAt, cancellation);
        var output = new MemoryStream();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            ZstandardArchive.Decompress(source, output, cancellationToken: cancellation.Token));

        Assert.Equal(0, output.Length);
        Assert.True(source.ReadsAfterCancellation <= 1);
    }

    [Fact]
    public void SingleSegmentFramesAreHeldToTheWindowLimit() {
        // zstd writes a known-size input smaller than its window as one segment, so the content size is the window.
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(Open("text-level1.zst"), new MemoryStream(), new ArchiveReadOptions { MaxWindowBytes = 64 * 1024 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxWindowBytes), error.LimitName);
    }

    [Fact]
    public void DictionariesLargerThanTheWindowLimitAreRefusedBeforeDecoding() {
        ZstandardDictionary dictionary = ZstandardDictionary.Load(ReadFixture("raw-content.dict"));
        var budget = new ArchiveOperationBudget();
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.DecompressWithDictionary(Open("message-raw-dict.zst"), new MemoryStream(), dictionary,
                new ArchiveReadOptions { MaxWindowBytes = 64, OperationBudget = budget }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxWindowBytes), error.LimitName);
        Assert.Equal(0, budget.Usage.InputBytes);
    }

    [Fact]
    public void ACancelledCallAdmitsNoInput() {
        var budget = new ArchiveOperationBudget();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() =>
            ZstandardArchive.Decompress(Open("empty.zst"), new MemoryStream(), new ArchiveReadOptions { OperationBudget = budget }, cancellation.Token));
        Assert.Equal(0, budget.Usage.InputBytes);
    }

    [Fact]
    public void CancellationWithTheLastHeaderByteChargesNoDictionaryWork() {
        byte[] frame = ReadFixture("message-raw-dict.zst");
        ZstandardDictionary dictionary = ZstandardDictionary.Load(ReadFixture("raw-content.dict"));
        var budget = new ArchiveOperationBudget();
        using var cancellation = new CancellationTokenSource();
        // The raw-dictionary frame's header ends at byte 6 (magic, descriptor and a 1-byte content size).
        var source = new TrickleStream(frame, 6, cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() => ZstandardArchive.DecompressWithDictionary(source, new MemoryStream(), dictionary,
            new ArchiveReadOptions { OperationBudget = budget }, cancellation.Token));

        Assert.Equal(0, budget.Usage.DecompressionWorkBytes);
    }

    [Theory]
    // Cancellation arrives with the last byte of a raw block header, or with the magic number of a skippable frame.
    [InlineData(false, 9)]
    [InlineData(true, 4)]
    public void CancellationBeforeABlockOrSkippableFrameChargesNoWork(bool skippable, long cancelAt) {
        var frame = new List<byte>(skippable
            ? new byte[] { 0x50, 0x2A, 0x4D, 0x18, 0x00, 0x00, 0x00, 0x00 }
            : new byte[] { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x00, 0x01, 0x20, 0x00 });
        if (!skippable) frame.AddRange(Enumerable.Repeat((byte)0x41, 1024));
        var budget = new ArchiveOperationBudget();
        using var cancellation = new CancellationTokenSource();
        var source = new TrickleStream(frame.ToArray(), cancelAt, cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() => ZstandardArchive.Decompress(source, new MemoryStream(),
            new ArchiveReadOptions { OperationBudget = budget }, cancellation.Token));

        Assert.Equal(0, budget.Usage.DecompressionWorkBytes);
    }

    [Fact]
    public void RawDictionariesShorterThanEightBytesAreRejected() {
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardDictionary.Load(new byte[] { 1, 2, 3, 4, 5, 6, 7 }));
        Assert.Equal(7, error.Offset);
    }

    [Fact]
    public void SkippableFramesCountAsDecodeWork() {
        byte[] empty = { 0x50, 0x2A, 0x4D, 0x18, 0x00, 0x00, 0x00, 0x00 };
        byte[] source = empty.Concat(empty).ToArray();
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream(),
                new ArchiveReadOptions { MaxDecompressionWorkBytes = 4096 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxDecompressionWorkBytes), error.LimitName);
    }

    [Fact]
    public void ASourceThatShrinksAfterAFrameIsReportedAsTruncated() {
        byte[] frame = ReadFixture("text-level1.zst");
        var source = new MemoryStream();
        source.Write(frame, 0, frame.Length);
        source.Write(frame, 0, frame.Length);
        source.Position = 0;
        var output = new TruncatingStream(source, frame.Length);

        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(source, output));

        Assert.Equal(frame.Length, error.Offset);
    }

    [Fact]
    public void ATruncatedTrailingSkippableFrameIsReported() {
        byte[] frame = ReadFixture("text-level1.zst");
        byte[] skippable = { 0x50, 0x2A, 0x4D, 0x18, 0x05, 0x00, 0x00, 0x00, 1, 2, 3, 4, 5 };
        var source = new MemoryStream();
        source.Write(frame, 0, frame.Length);
        source.Write(skippable, 0, skippable.Length);
        source.Position = 0;

        Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(source, new TruncatingStream(source, frame.Length + 10)));
    }

    [Fact]
    public void ContentSizesBeyondInt64AreRefusedAsOutput() {
        // Single segment with an 8-byte content size of 2^63.
        byte[] source = { 0x28, 0xB5, 0x2F, 0xFD, 0xE0, 0, 0, 0, 0, 0, 0, 0, 0x80 };
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream()));
        Assert.Equal(nameof(ArchiveReadOptions.MaxEntryBytes), error.LimitName);
    }

    [Fact]
    public void GrowingTheHistoryChargesTheWholeBufferCopy() {
        // A 128 KiB window and RLE blocks of 128 KiB, 1 byte and 128 KiB. The third block grows the 256 KiB buffer, which
        // copies all of it: blocks 131,073 + 4,096 + 131,073 plus the 262,144-byte copy make 528,386.
        byte[] source = {
            0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x38,
            0x02, 0x00, 0x10, 0x41,
            0x0A, 0x00, 0x00, 0x42,
            0x03, 0x00, 0x10, 0x43,
        };
        Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream(),
                new ArchiveReadOptions { MaxDecompressionWorkBytes = 528385 }));
        var output = new MemoryStream();
        ZstandardArchive.Decompress(new MemoryStream(source), output, new ArchiveReadOptions { MaxDecompressionWorkBytes = 528386 });
        Assert.Equal(2 * 128 * 1024 + 1, output.Length);
    }

    [Fact]
    public void BlockWorkIsChargedBeforeTheBlockIsDecoded() {
        // A 1 KiB window and a last compressed block of four invalid bytes: charging comes first, so the work limit is
        // reported rather than the corruption decoding would find.
        byte[] source = { 0x28, 0xB5, 0x2F, 0xFD, 0x00, 0x00, 0x25, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF };
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream(),
                new ArchiveReadOptions { MaxDecompressionWorkBytes = 4095 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxDecompressionWorkBytes), error.LimitName);
        Assert.Throws<ArchiveCorruptionException>(() =>
            ZstandardArchive.Decompress(new MemoryStream(source), new MemoryStream()));
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

    /// <summary>Output that truncates the source once the first frame's output arrives, as if the file shrank mid-read.</summary>
    private sealed class TruncatingStream : MemoryStream {
        private readonly MemoryStream _source;
        private readonly long _length;
        public TruncatingStream(MemoryStream source, long length) {
            _source = source;
            _length = length;
        }
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            if (_source.Length > _length) _source.SetLength(_length);
        }
    }

    /// <summary>Returns one byte per read and cancels once the byte at a given offset has been read.</summary>
    private sealed class TrickleStream : MemoryStream {
        private readonly long _cancelAt;
        private readonly CancellationTokenSource _cancellation;
        public TrickleStream(byte[] data, long cancelAt, CancellationTokenSource cancellation) : base(data, writable: false) {
            _cancelAt = cancelAt;
            _cancellation = cancellation;
        }
        public int ReadsAfterCancellation { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_cancellation.IsCancellationRequested) ReadsAfterCancellation++;
            int read = base.Read(buffer, offset, Math.Min(count, 1));
            if (Position >= _cancelAt) _cancellation.Cancel();
            return read;
        }
    }
}
