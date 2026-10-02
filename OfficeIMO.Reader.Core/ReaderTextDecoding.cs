using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Applies explicit decoding policy after detecting the Unicode BOM without buffering the document.</summary>
internal sealed class ReaderTextDecoding {
    internal int InvalidSequences { get; private set; }

    internal StreamReader Open(Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        byte[] prefix = new byte[4];
        int count = 0;
        while (count < prefix.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(prefix, count, prefix.Length - count);
            if (read == 0) break;
            count += read;
        }
        Encoding encoding = options.TextEncoding ?? Encoding.UTF8;
        int skip = 0;
        if (count >= 4 && prefix[0] == 0 && prefix[1] == 0 && prefix[2] == 0xfe && prefix[3] == 0xff) {
            encoding = new UTF32Encoding(true, true); skip = 4;
        } else if (count >= 4 && prefix[0] == 0xff && prefix[1] == 0xfe && prefix[2] == 0 && prefix[3] == 0) {
            encoding = new UTF32Encoding(false, true); skip = 4;
        } else if (count >= 3 && prefix[0] == 0xef && prefix[1] == 0xbb && prefix[2] == 0xbf) {
            encoding = Encoding.UTF8; skip = 3;
        } else if (count >= 2 && prefix[0] == 0xfe && prefix[1] == 0xff) {
            encoding = Encoding.BigEndianUnicode; skip = 2;
        } else if (count >= 2 && prefix[0] == 0xff && prefix[1] == 0xfe) {
            encoding = Encoding.Unicode; skip = 2;
        }
        encoding = (Encoding)encoding.Clone();
        encoding.DecoderFallback = options.ThrowOnInvalidTextBytes
            ? DecoderFallback.ExceptionFallback
            : new DiagnosticFallback(this);
        return new StreamReader(new ReaderPrefixStream(stream, prefix, skip, count), encoding,
            detectEncodingFromByteOrderMarks: false, bufferSize: 4096, leaveOpen: false);
    }

    private sealed class DiagnosticFallback : DecoderFallback {
        private readonly ReaderTextDecoding _owner;
        internal DiagnosticFallback(ReaderTextDecoding owner) { _owner = owner; }
        public override int MaxCharCount => 1;
        public override DecoderFallbackBuffer CreateFallbackBuffer() => new Buffer(_owner);
        private sealed class Buffer : DecoderFallbackBuffer {
            private readonly ReaderTextDecoding _owner;
            private bool _remaining;
            internal Buffer(ReaderTextDecoding owner) { _owner = owner; }
            public override bool Fallback(byte[] bytesUnknown, int index) {
                _owner.InvalidSequences++;
                _remaining = true;
                return true;
            }
            public override char GetNextChar() {
                if (!_remaining) return '\0';
                _remaining = false;
                return '\ufffd';
            }
            public override bool MovePrevious() { if (_remaining) return false; _remaining = true; return true; }
            public override int Remaining => _remaining ? 1 : 0;
            public override void Reset() { _remaining = false; base.Reset(); }
        }
    }

}
