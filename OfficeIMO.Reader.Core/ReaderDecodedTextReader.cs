using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Feeds the configured decoder and explicitly flushes incomplete sequences at end of input.</summary>
/// <remarks>Used by the plain-text adapter so replacement diagnostics and strict decoding have the same
/// contract on every runtime. Disposing it closes the replay wrapper, which leaves the caller's stream open.</remarks>
internal sealed class ReaderDecodedTextReader : TextReader {
    private readonly Stream _stream;
    private readonly Decoder _decoder;
    private readonly CancellationToken _cancellationToken;
    private readonly byte[] _bytes = new byte[4096];
    private readonly char[] _characters;
    private int _position;
    private int _count;
    private bool _endOfInput;
    private bool _disposed;

    internal ReaderDecodedTextReader(Stream stream, Encoding encoding, CancellationToken cancellationToken) {
        _stream = stream;
        _decoder = encoding.GetDecoder();
        _characters = new char[encoding.GetMaxCharCount(_bytes.Length)];
        _cancellationToken = cancellationToken;
    }

    public override int Read() {
        if (_disposed) throw new ObjectDisposedException(nameof(ReaderDecodedTextReader));
        while (_position == _count) {
            if (_endOfInput) return -1;
            _cancellationToken.ThrowIfCancellationRequested();
            int bytesRead = _stream.Read(_bytes, 0, _bytes.Length);
            _endOfInput = bytesRead == 0;
            _count = _decoder.GetChars(_bytes, 0, bytesRead, _characters, 0, flush: _endOfInput);
            _position = 0;
        }
        return _characters[_position++];
    }

    protected override void Dispose(bool disposing) {
        if (disposing && !_disposed) {
            _disposed = true;
            _stream.Dispose();
        }
        base.Dispose(disposing);
    }
}
