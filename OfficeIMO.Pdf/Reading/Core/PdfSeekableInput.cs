using OfficeIMO.Core.Internal;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Pdf;

/// <summary>Accesses one stable PDF window without taking ownership of the caller's stream.</summary>
internal sealed class PdfSeekableInput : IDisposable {
    private readonly SeekableReadWindowStream _window;
    private readonly Stream _source;
    private readonly bool _ownsSource;
    private bool _disposed;
    [ThreadStatic] private static Admission? _activeAdmission;

    internal PdfSeekableInput(Stream source, long offset, long length, PdfReadLimits limits, bool ownsSource = false) {
        Guard.NotNull(limits, nameof(limits));
        limits.Validate();
        long limit = Math.Min(limits.MaxInputBytes, int.MaxValue);
        if (length > limit) throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limit, length);
        _window = new SeekableReadWindowStream(source, offset, length, preserveSourcePosition: true);
        _source = source;
        _ownsSource = ownsSource;
        Length = (int)length;
    }

    internal int Length { get; }

    internal bool IsSource(Stream stream) => ReferenceEquals(_source, stream);

    internal Admission BeginAdmission() => new(this);

    internal byte[] ReadAt(int offset, int length, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRange(offset, length);
        byte[] buffer = new byte[length];
        ReadAt(offset, buffer, 0, length, cancellationToken);
        return buffer;
    }

    internal void ReadAt(int offset, byte[] buffer, int bufferOffset, int length, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(buffer, nameof(buffer));
        ValidateRange(offset, length);
        if (bufferOffset < 0 || bufferOffset > buffer.Length - length) {
            throw new ArgumentOutOfRangeException(nameof(bufferOffset));
        }

        while (!Monitor.TryEnter(_window, millisecondsTimeout: 25)) {
            cancellationToken.ThrowIfCancellationRequested();
        }
        try {
#if NET8_0_OR_GREATER
            ObjectDisposedException.ThrowIf(_disposed, this);
#else
            if (_disposed) throw new ObjectDisposedException(nameof(PdfSeekableInput));
#endif
            Admission? admission = _activeAdmission?.Source == this ? _activeAdmission : null;
            admission?.ThrowIfReadFailed();
            try {
                _window.Position = offset;
                while (length > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = _window.Read(buffer, bufferOffset, Math.Min(length, 65536));
                    if (read == 0) throw new EndOfStreamException("The PDF input ended before its declared length.");
                    bufferOffset += read;
                    length -= read;
                }
            } catch (Exception exception) when (exception is not (OperationCanceledException or OutOfMemoryException or StackOverflowException)) {
                admission?.CaptureReadFailure(exception);
                throw;
            }
        } finally {
            Monitor.Exit(_window);
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    public void Dispose() {
        lock (_window) {
            if (_disposed) return;
            _disposed = true;
            _window.Dispose();
            if (_ownsSource) _source.Dispose();
        }
    }

    internal void CopyTo(Stream output, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var buffer = new byte[Math.Min(Length, 65536)];
        for (int offset = 0; offset < Length;) {
            int count = Math.Min(buffer.Length, Length - offset);
            ReadAt(offset, buffer, 0, count, cancellationToken);
            output.Write(buffer, 0, count);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal void ValidateRange(int offset, int length) {
        if (length < 0 || offset < 0 || offset > Length - length) {
            throw new ArgumentOutOfRangeException(nameof(offset), "The read must stay inside the PDF window.");
        }
    }
    // The canonical parser is synchronous. Its I/O fault belongs to this admission,
    // so an unrelated caller can still retry a deferred read after a transient fault.
    internal sealed class Admission : IDisposable {
        private readonly Admission? _previous;
        private ExceptionDispatchInfo? _failure;

        internal Admission(PdfSeekableInput source) {
            Source = source;
            _previous = _activeAdmission;
            _activeAdmission = this;
        }

        internal PdfSeekableInput Source { get; }

        internal void CaptureReadFailure(Exception failure) => _failure ??= ExceptionDispatchInfo.Capture(failure);

        internal void ThrowIfReadFailed() => _failure?.Throw();

        public void Dispose() => _activeAdmission = _previous;
    }

}
