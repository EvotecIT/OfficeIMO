using System.Text;
using OfficeIMO.Toml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// Stream and TextReader input go through the pooled sliding buffer, which the in-memory suites never exercise. These
// tests compare buffered reading, with tiny buffers and one byte or character per read, against in-memory reading of
// the same text: every token, position and error must be the same wherever the buffer boundaries fall.
public class TomlReaderInputTests {
    private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);
    private static readonly TomlReaderOptions Comments = new TomlReaderOptions { IncludeComments = true };

    [Fact]
    public void SuiteAndCorpusReadTheSameThroughTinyBuffers() {
        var inputs = TomlTestSuite.Valid.Concat(TomlTestSuite.Invalid).Select(c => (c.Name, c.Toml))
            .Concat(TomlCorpus.Load().Select(f => (f.Path, f.Bytes)));
        foreach ((string name, byte[] bytes) in inputs) {
            string? text = TryDecode(bytes);
            if (text is null) {
                // Invalid UTF-8: the stream reader fails at the first bad byte, where loading the whole file fails too,
                // unless a syntax error comes first.
                TomlException loaded = Assert.ThrowsAny<TomlException>(() => TomlDocument.Load(new MemoryStream(bytes)));
                string streamed = Trace(() => new TomlReader(new TomlStreamSource(new TrickleStream(bytes, 1)), Comments, default, 2));
                if (streamed.Contains(TomlErrorCodes.InvalidEncoding)) Assert.EndsWith(" " + loaded.Position, streamed);
                else Assert.Contains("TomlException", streamed);
                continue;
            }
            string expected = Trace(() => new TomlReader(text, Comments));
            foreach (int size in new[] { 2, 3, 7, 64 }) {
                Assert.True(expected == Trace(() => new TomlReader(new TomlStreamSource(new TrickleStream(bytes, 1)), Comments, default, size)), name + " differs through a " + size + "-character buffer from a stream.");
                Assert.True(expected == Trace(() => new TomlReader(new TomlTextReaderSource(new TrickleReader(text, 1)), Comments, default, size)), name + " differs through a " + size + "-character buffer from a text reader.");
            }
        }
    }

    // A long input makes the buffer compact between tokens and grow under multi-line strings and long lines, with CRLF
    // line breaks, comments and multi-byte characters at every offset.
    [Fact]
    public void LargeInputCompactsAndGrowsTheBufferWithoutChangingTokens() {
        var text = new StringBuilder();
        for (int i = 0; i < 600; i++) {
            string newline = i % 3 == 0 ? "\r\n" : "\n";
            text.Append("[[package]]").Append(newline);
            text.Append("name = \"crate-").Append(i).Append("\" # é ").Append("😀").Append(newline);
            text.Append("deps = [").Append(newline).Append("  { name = 'a', version = \"1.0\" },").Append(newline).Append("]").Append(newline);
            text.Append("notes = '''").Append(newline);
            int lines = i % 100 == 0 ? 3000 : 2;
            for (int line = 0; line < lines; line++) text.Append("  line ").Append(line).Append(newline);
            text.Append("'''").Append(newline);
            text.Append("escaped = \"\"\"\\").Append(newline).Append("   continued \\u00e9\"\"\"").Append(newline);
        }
        text.Append("long = \"").Append('x', 70_000).Append("\"\n");
        string toml = text.ToString();
        byte[] bytes = Encoding.UTF8.GetBytes(toml);
        string expected = Trace(() => new TomlReader(toml, Comments));
        Assert.Equal(expected, Trace(() => new TomlReader(new MemoryStream(bytes), Comments)));
        Assert.Equal(expected, Trace(() => new TomlReader(new TrickleReader(toml, 4093), Comments)));
        Assert.Equal(expected, Trace(() => new TomlReader(new TomlStreamSource(new TrickleStream(bytes, 509)), Comments, default, 61)));
    }

    [Theory]
    [InlineData(new byte[] { (byte)'a', (byte)' ', (byte)'=', (byte)' ', (byte)'"', 0xC3, 0x28, (byte)'"', (byte)'\n' }, 1, 6)]
    [InlineData(new byte[] { (byte)'a', (byte)'=', (byte)'1', (byte)'\n', (byte)'#', 0xE2, 0x82 }, 2, 2)]
    [InlineData(new byte[] { 0xEF, 0xBB, 0xBF, (byte)'#', 0xF0, 0x9F, 0x98, 0x80, 0xED, 0xA0, 0x80, (byte)'\n' }, 1, 4)]
    public void InvalidUtf8FailsAtTheFirstBadByte(byte[] bytes, int line, int column) {
        using var reader = new TomlReader(new TrickleStream(bytes, 1));
        TomlException error = Assert.ThrowsAny<TomlException>(() => { while (reader.Read()) { } });
        Assert.Equal(TomlErrorCodes.InvalidEncoding, error.Code);
        Assert.Equal((line, column), (error.Line, error.Column));
    }

    [Fact]
    public void BufferedInputEnforcesInputAndStringLimits() {
        var tooLong = new TomlReaderOptions { MaxInputBytes = 10 };
        using (var reader = new TomlReader(new MemoryStream(Encoding.UTF8.GetBytes("a = 1\nb = 2\n")), tooLong)) {
            TomlLimitExceededException error = Assert.Throws<TomlLimitExceededException>(() => { while (reader.Read()) { } });
            Assert.Equal(nameof(TomlReaderOptions.MaxInputBytes), error.LimitName);
        }
        // An unterminated multi-line string stops growing the buffer once it passes the limit.
        var shortStrings = new TomlReaderOptions { MaxStringLength = 100 };
        using (var reader = new TomlReader(new TrickleReader("s = '''\n" + string.Concat(Enumerable.Repeat("text\n", 100_000)), 64), shortStrings)) {
            Assert.True(reader.Read());
            TomlLimitExceededException error = Assert.Throws<TomlLimitExceededException>(() => reader.Read());
            Assert.Equal(nameof(TomlReaderOptions.MaxStringLength), error.LimitName);
            Assert.Equal((1, 5), (error.Line, error.Column));
        }
    }

    [Fact]
    public void DisposingTheReaderLeavesTheCallersStreamOpen() {
        var stream = new MemoryStream(Encoding.UTF8.GetBytes("a = 1\n"));
        var reader = new TomlReader(stream);
        Assert.True(reader.Read());
        reader.Dispose();
        Assert.True(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => reader.Read());
    }

    private static string? TryDecode(byte[] bytes) {
        try {
            int offset = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
            return StrictUtf8.GetString(bytes, offset, bytes.Length - offset);
        } catch (DecoderFallbackException) {
            return null;
        }
    }

    /// <summary>Every token with its text, positions, depth and style, then the end or the error.</summary>
    private static string Trace(Func<TomlReader> create) {
        var trace = new StringBuilder();
        try {
            using TomlReader reader = create();
            while (reader.Read()) {
                trace.Append(reader.TokenType).Append(' ').Append(reader.ValueSpan.ToString()).Append(' ')
                    .Append(reader.Start).Append('-').Append(reader.End).Append(' ')
                    .Append(reader.Start.Offset).Append('-').Append(reader.End.Offset).Append(' ')
                    .Append(reader.Depth).Append(' ').Append(reader.StringStyle).Append('\n');
            }
            trace.Append("end");
        } catch (TomlException ex) {
            trace.Append(ex.GetType().Name).Append(' ').Append(ex.Code).Append(' ').Append(ex.Message).Append(' ').Append(ex.Position.Offset).Append(' ').Append(ex.Position);
        }
        return trace.ToString();
    }

    /// <summary>A stream that returns at most a few bytes per call, splitting multi-byte characters across reads.</summary>
    private sealed class TrickleStream : Stream {
        private readonly byte[] _bytes;
        private readonly int _chunk;
        private int _position;

        public TrickleStream(byte[] bytes, int chunk) {
            _bytes = bytes;
            _chunk = chunk;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _position; set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count) {
            int n = Math.Min(Math.Min(count, _chunk), _bytes.Length - _position);
            Array.Copy(_bytes, _position, buffer, offset, n);
            _position += n;
            return n;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    /// <summary>A text reader that returns at most a few characters per call.</summary>
    private sealed class TrickleReader : TextReader {
        private readonly string _text;
        private readonly int _chunk;
        private int _position;

        public TrickleReader(string text, int chunk) {
            _text = text;
            _chunk = chunk;
        }

        public override int Read(char[] buffer, int index, int count) {
            int n = Math.Min(Math.Min(count, _chunk), _text.Length - _position);
            _text.CopyTo(_position, buffer, index, n);
            _position += n;
            return n;
        }
    }
}
