#if NET8_0_OR_GREATER
using System.Text;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(7, 16, false)]
        [InlineData(7, 16, true)]
        [InlineData(31, 16, false)]
        [InlineData(31, 16, true)]
        [InlineData(4095, 8192, false)]
        [InlineData(4095, 8192, true)]
        [InlineData(4096, 16, false)]
        [InlineData(4096, 16, true)]
        public void PooledUtf8TextWriter_AsciiPrefixPreservesRawAndFollowingCharacterOrder(int prefixLength, int bufferSize, bool emptyBytes) {
            string prefix = new('a', prefixLength);
            byte[] raw = emptyBytes ? [] : " byte 🚀 "u8.ToArray();
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), bufferSize, leaveOpen: true)) {
                writer.Write(prefix);
                writer.WriteUtf8(raw);
                writer.Write('λ');
                writer.Write(" suffix");
            }
            Assert.Equal(Encoding.UTF8.GetBytes(prefix).Concat(raw).Concat(Encoding.UTF8.GetBytes("λ suffix")), output.ToArray());
        }

        [Theory]
        [InlineData(false, false, false, false)]
        [InlineData(false, false, false, true)]
        [InlineData(false, false, true, false)]
        [InlineData(false, false, true, true)]
        [InlineData(false, true, false, false)]
        [InlineData(false, true, false, true)]
        [InlineData(false, true, true, false)]
        [InlineData(false, true, true, true)]
        [InlineData(true, false, false, false)]
        [InlineData(true, false, false, true)]
        [InlineData(true, false, true, false)]
        [InlineData(true, false, true, true)]
        [InlineData(true, true, false, false)]
        [InlineData(true, true, false, true)]
        [InlineData(true, true, true, false)]
        [InlineData(true, true, true, true)]
        public void PooledUtf8TextWriter_RawBoundaryCompletesBufferedOrPendingSurrogates(bool longInput, bool flushBeforeBytes, bool completePair, bool emptyBytes) {
            string before = (longInput ? new string('a', 4096) : "prefix ") + (completePair ? "\ud83d\ude80" : "\ud83d");
            byte[] raw = emptyBytes ? [] : " raw 🚀 "u8.ToArray();
            const string after = "\ude80 suffix λ";
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 16, leaveOpen: true)) {
                writer.Write(before);
                if (flushBeforeBytes) writer.Flush();
                writer.WriteUtf8(raw);
                writer.Write(after);
            }
            // Even an empty raw write ends the preceding UTF-16 value, so the
            // following low surrogate cannot complete its isolated high surrogate.
            Assert.Equal(Encoding.UTF8.GetBytes(before).Concat(raw).Concat(Encoding.UTF8.GetBytes(after)), output.ToArray());
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void PooledUtf8TextWriter_StrictFallbackRejectsSurrogatesBeforeRawBytes(bool pending, bool emptyBytes) {
            using var output = new MemoryStream();
            var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false, true), 16, leaveOpen: true);
            try {
                writer.Write("prefix \ud83d");
                if (pending) writer.Flush();
                byte[] raw = emptyBytes ? [] : "RAW"u8.ToArray();
                Assert.Throws<EncoderFallbackException>(() => writer.WriteUtf8(raw));
            } finally {
                // Disposal can retry the invalid text. Preserve the boundary
                // error while releasing the writer through its normal cleanup.
                try { writer.Dispose(); } catch (EncoderFallbackException) { }
            }
            Assert.DoesNotContain("RAW", Encoding.UTF8.GetString(output.ToArray()));
        }

        [Fact]
        public void PooledUtf8TextWriter_CustomUtf8EncoderRetainsCharacterFlushCall() {
            var encoding = new RecordingUtf8Encoding();
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, encoding, 16, leaveOpen: true)) {
                writer.Write("prefix");
                writer.WriteUtf8("RAW"u8);
                Assert.Equal(new[] { (6, true) }, encoding.Calls);
                writer.Write("suffix");
            }
            Assert.Equal("prefixRAWsuffix"u8.ToArray(), output.ToArray());
        }

        [Fact]
        public void PooledUtf8TextWriter_CustomFallbackRetainsBufferedCharacterIndex() {
            var encoding = (Encoding)new UTF8Encoding(false).Clone();
            encoding.EncoderFallback = new IndexEncoderFallback();
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, encoding, 32, leaveOpen: true)) {
                writer.Write("prefix\ud83d");
                writer.WriteUtf8("RAW"u8);
                writer.Write("suffix");
            }
            Assert.Equal("prefix6RAWsuffix"u8.ToArray(), output.ToArray());
        }

        [Fact]
        public void PooledUtf8TextWriter_NonUtf8EncodingRetainsMixedByteOrder() {
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, Encoding.Unicode, 16, leaveOpen: true)) {
                writer.Write("prefix λ");
                writer.WriteUtf8("RAW"u8);
                writer.Write("suffix");
            }
            Assert.Equal(Encoding.Unicode.GetBytes("prefix λ").Concat("RAW"u8.ToArray()).Concat(Encoding.Unicode.GetBytes("suffix")), output.ToArray());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PooledUtf8TextWriter_DisposedRejectsEmptyAndNonemptyRawWrites(bool emptyBytes) {
            using var output = new MemoryStream();
            var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 16, leaveOpen: true);
            writer.Dispose();
            byte[] raw = emptyBytes ? [] : "RAW"u8.ToArray();
            Assert.Throws<ObjectDisposedException>(() => writer.WriteUtf8(raw));
            Assert.Empty(output.ToArray());
        }

        private sealed class RecordingUtf8Encoding : UTF8Encoding {
            internal List<(int CharacterCount, bool Flush)> Calls { get; } = [];
            public override Encoder GetEncoder() => new RecordingEncoder(base.GetEncoder(), Calls);
        }

        private sealed class RecordingEncoder(Encoder inner, List<(int CharacterCount, bool Flush)> calls) : Encoder {
            public override int GetByteCount(char[] chars, int index, int count, bool flush) => inner.GetByteCount(chars, index, count, flush);
            public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex, bool flush) =>
                inner.GetBytes(chars, charIndex, charCount, bytes, byteIndex, flush);
            public override void Convert(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex, int byteCount,
                bool flush, out int charsUsed, out int bytesUsed, out bool completed) {
                calls.Add((charCount, flush));
                inner.Convert(chars, charIndex, charCount, bytes, byteIndex, byteCount, flush, out charsUsed, out bytesUsed, out completed);
            }
        }

        private sealed class IndexEncoderFallback : EncoderFallback {
            public override int MaxCharCount => 1;
            public override EncoderFallbackBuffer CreateFallbackBuffer() => new IndexEncoderFallbackBuffer();
        }

        private sealed class IndexEncoderFallbackBuffer : EncoderFallbackBuffer {
            private char _replacement;
            private bool _remaining;
            public override int Remaining => _remaining ? 1 : 0;
            public override bool Fallback(char charUnknown, int index) {
                _replacement = index < 0 ? '!' : (char)('0' + index % 10);
                _remaining = true;
                return true;
            }
            public override bool Fallback(char charUnknownHigh, char charUnknownLow, int index) => Fallback(charUnknownHigh, index);
            public override char GetNextChar() {
                if (!_remaining) return '\0';
                _remaining = false;
                return _replacement;
            }
            public override bool MovePrevious() {
                if (_remaining) return false;
                _remaining = true;
                return true;
            }
            public override void Reset() => _remaining = false;
        }
    }
}
#endif
