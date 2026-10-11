using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherInputContractTests {
    [Fact]
    public void Caller_stream_reads_from_its_current_position_and_remains_open() {
        byte[] input = File.ReadAllBytes(PublisherNativeTests.Fixture("Simple.pub"));
        using var stream = new MemoryStream(new byte[7].Concat(input).ToArray());
        stream.Position = 7;
        PublisherDocument document = PublisherDocument.Load(stream);
        Assert.Single(document.Pages);
        Assert.True(stream.CanRead);
        Assert.Equal(stream.Length, stream.Position);
        using var nonSeekable = new NonSeekableStream(input);
        Assert.Single(PublisherDocument.Load(nonSeekable).Pages);
        Assert.False(nonSeekable.Disposed);
    }

    [Fact]
    public void Cancellation_is_observed_before_input_access_and_during_non_seekable_reads() {
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => PublisherDocument.Load("missing.pub", cancellationToken: cancelled.Token));
        using var source = new CancellationTokenSource();
        using var stream = new NonSeekableStream(File.ReadAllBytes(PublisherNativeTests.Fixture("Simple.pub")), source);
        Assert.Throws<OperationCanceledException>(() => PublisherDocument.Load(stream, cancellationToken: source.Token));
        Assert.False(stream.Disposed);
    }

    [Fact]
    public void Resource_limits_reject_whole_input_instead_of_returning_partial_documents() {
        string file = PublisherNativeTests.Fixture("Sample.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions { MaximumPages = 1 }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxInputBytes = 100 } }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxRecords = 100 } }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 100 } }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(PublisherNativeTests.Fixture("SampleBrochure.pub"), new PublisherReadOptions { MaximumImageBytes = 100 }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(PublisherNativeTests.Fixture("SampleBrochure.pub"), new PublisherReadOptions { MaximumTotalImageBytes = 40000 }));
    }

    [Theory]
    [InlineData("Contents", 0x1A)]
    [InlineData("Escher/EscherStm", 4)]
    public void Out_of_range_native_offsets_are_rejected(string stream, int offset) {
        byte[] input = Mutate(stream, bytes => WriteUInt32(bytes, offset, uint.MaxValue));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    [Fact]
    public void Cyclic_quill_directories_are_rejected() {
        byte[] input = Mutate("Quill/QuillSub/CONTENTS", bytes => WriteUInt32(bytes, 0x1C, 0x18));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    [Fact]
    public void Native_text_requires_well_formed_utf16() {
        byte[] input = Mutate("Quill/QuillSub/CONTENTS", bytes => { bytes[0x200] = 0; bytes[0x201] = 0xD8; });
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    [Fact]
    public void Projected_repetition_is_bounded_in_addition_to_source_content() {
        string file = PublisherNativeTests.Fixture("Sample.pub");
        PublisherDocument baseline = PublisherDocument.Load(file);
        long sourceCharacters = baseline.TextStories.Sum(story => (long)story.Text.Length);
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxTextCharacters = (int)sourceCharacters + 1 }
        }));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 8 }
        }));
    }

    internal static byte[] Mutate(string name, Action<byte[]> mutate, string file = "Simple.pub") {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(PublisherNativeTests.Fixture(file)), out OfficeCompoundFile? source, out string? error), error);
        Assert.True(source!.Streams.TryGetValue(name, out byte[]? bytes));
        bytes = (byte[])bytes!.Clone(); mutate(bytes);
        return OfficeCompoundFileWriter.Rewrite(source!, new Dictionary<string, byte[]> { [name] = bytes });
    }
    internal static void WriteUInt32(byte[] bytes, int offset, uint value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8));
    }
    private sealed class NonSeekableStream : Stream {
        private readonly MemoryStream _input;
        private readonly CancellationTokenSource? _cancel;
        internal NonSeekableStream(byte[] bytes, CancellationTokenSource? cancel = null) { _input = new MemoryStream(bytes); _cancel = cancel; }
        internal bool Disposed { get; private set; }
        public override bool CanRead => !Disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) { int read = _input.Read(buffer, offset, Math.Min(count, 64)); _cancel?.Cancel(); return read; }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { Disposed = true; if (disposing) _input.Dispose(); base.Dispose(disposing); }
    }
}
