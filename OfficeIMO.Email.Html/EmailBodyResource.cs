namespace OfficeIMO.Email;

/// <summary>Operation-scoped attachment resource resolved by CID, content location, or filename.</summary>
public sealed class EmailBodyResource {
    private readonly EmailAttachment _attachment;
    private readonly long _maximumBytes;
    private readonly EmailBodyResourceBudget _budget;
    private readonly EmailBodyResourceLimitState _limitState;

    internal EmailBodyResource(
        EmailAttachment attachment,
        long maximumBytes,
        EmailBodyResourceBudget budget) {
        _attachment = attachment ?? throw new ArgumentNullException(nameof(attachment));
        _maximumBytes = maximumBytes;
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
        _limitState = new EmailBodyResourceLimitState(maximumBytes);
    }

    /// <summary>Content type declared by the artifact.</summary>
    public string ContentType => _attachment.ContentType ?? "application/octet-stream";
    /// <summary>Declared decoded length.</summary>
    public long Length => _attachment.Length;
    /// <summary>Normalized Content-ID without angle brackets.</summary>
    public string? ContentId => NormalizeContentId(_attachment.ContentId);
    /// <summary>Content-Location retained by the artifact.</summary>
    public string? ContentLocation => _attachment.ContentLocation;
    /// <summary>Safe filename retained by the artifact.</summary>
    public string? FileName => _attachment.FileName;

    /// <summary>Opens a fresh sequential stream governed by the per-resource and projection-wide budgets.</summary>
    public Stream OpenReadStream(CancellationToken cancellationToken = default) {
        EnsureDeclaredLengthAllowed();
        _limitState.ThrowIfExceeded();
        _budget.ThrowIfExceeded();
        cancellationToken.ThrowIfCancellationRequested();
        Stream source = _attachment.OpenContentStream();
        try {
            return new EmailBodyResourceReadStream(
                source,
                _maximumBytes,
                _budget,
                _limitState,
                cancellationToken);
        } catch {
            source.Dispose();
            throw;
        }
    }

    /// <summary>Asynchronously opens a fresh sequential stream governed by the configured budgets.</summary>
    public async Task<Stream> OpenReadStreamAsync(CancellationToken cancellationToken = default) {
        EnsureDeclaredLengthAllowed();
        _limitState.ThrowIfExceeded();
        _budget.ThrowIfExceeded();
        cancellationToken.ThrowIfCancellationRequested();
        Stream source = await _attachment.OpenContentStreamAsync(cancellationToken).ConfigureAwait(false);
        try {
            return new EmailBodyResourceReadStream(
                source,
                _maximumBytes,
                _budget,
                _limitState,
                cancellationToken);
        } catch {
            source.Dispose();
            throw;
        }
    }

    /// <summary>Copies this resource to a caller-owned destination without materializing another byte array.</summary>
    public void CopyTo(Stream destination, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        using Stream source = OpenReadStream(cancellationToken);
        var buffer = new byte[64 * 1024];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            destination.Write(buffer, 0, read);
        }
    }

    /// <summary>Asynchronously copies this resource to a caller-owned destination.</summary>
    public async Task CopyToAsync(Stream destination, CancellationToken cancellationToken = default) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        using Stream source = await OpenReadStreamAsync(cancellationToken).ConfigureAwait(false);
        var buffer = new byte[64 * 1024];
        while (true) {
            int read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            await destination.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Reads this resource within the configured budgets.</summary>
    public byte[] ReadAllBytes(CancellationToken cancellationToken = default) {
        using var output = CreateOutputBuffer();
        CopyTo(output, cancellationToken);
        return output.ToArray();
    }

    /// <summary>Asynchronously reads this resource within the configured budgets.</summary>
    public async Task<byte[]> ReadAllBytesAsync(CancellationToken cancellationToken = default) {
        using var output = CreateOutputBuffer();
        await CopyToAsync(output, cancellationToken).ConfigureAwait(false);
        return output.ToArray();
    }

    private MemoryStream CreateOutputBuffer() {
        EnsureDeclaredLengthAllowed();
        int capacity = Length > 0 && Length <= int.MaxValue ? checked((int)Length) : 0;
        return capacity > 0 ? new MemoryStream(capacity) : new MemoryStream();
    }

    private void EnsureDeclaredLengthAllowed() {
        if (Length > _maximumBytes) {
            throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                Length, _maximumBytes);
        }
    }

    internal static string? NormalizeContentId(string? value) => string.IsNullOrWhiteSpace(value)
        ? null
        : value!.Trim().Trim('<', '>');
}

internal sealed class EmailBodyResourceBudget {
    private readonly object _gate = new object();
    private readonly long _maximumBytes;
    private long _consumedBytes;
    private long _reservedBytes;
    private long? _exceededActualValue;

    internal EmailBodyResourceBudget(long maximumBytes) {
        _maximumBytes = maximumBytes;
    }

    internal int Reserve(int requestedCount) {
        if (requestedCount <= 0) return 0;
        lock (_gate) {
            ThrowIfExceededLocked();
            long available = _maximumBytes - _consumedBytes - _reservedBytes;
            if (available <= 0) return 0;
            int reserved = (int)Math.Min(requestedCount, available);
            _reservedBytes += reserved;
            return reserved;
        }
    }

    internal void Commit(int reservedCount, int consumedCount) {
        lock (_gate) {
            _reservedBytes -= reservedCount;
            _consumedBytes += consumedCount;
        }
    }

    internal void Release(int reservedCount) {
        if (reservedCount <= 0) return;
        lock (_gate) {
            _reservedBytes -= reservedCount;
        }
    }

    internal void MarkExceeded() {
        lock (_gate) {
            _exceededActualValue ??= _maximumBytes == long.MaxValue
                ? long.MaxValue
                : _maximumBytes + 1L;
        }
    }

    internal void ThrowIfExceeded() {
        lock (_gate) {
            ThrowIfExceededLocked();
        }
    }

    private void ThrowIfExceededLocked() {
        if (_exceededActualValue.HasValue) {
            throw new EmailLimitExceededException(
                "EmailBodyProjectionOptions.MaxTotalResourceBytes",
                _exceededActualValue.Value,
                _maximumBytes);
        }
    }
}

internal sealed class EmailBodyResourceLimitState {
    private readonly object _gate = new object();
    private readonly long _maximumBytes;
    private long? _exceededActualValue;

    internal EmailBodyResourceLimitState(long maximumBytes) {
        _maximumBytes = maximumBytes;
    }

    internal void MarkExceeded(long actualValue) {
        lock (_gate) {
            _exceededActualValue ??= actualValue;
        }
    }

    internal void ThrowIfExceeded() {
        lock (_gate) {
            if (_exceededActualValue.HasValue) {
                throw new EmailLimitExceededException(
                    "EmailBodyProjectionOptions.MaxResourceBytes",
                    _exceededActualValue.Value,
                    _maximumBytes);
            }
        }
    }
}

internal sealed class EmailBodyResourceReadStream : Stream {
    private readonly Stream _source;
    private readonly long _maximumBytes;
    private readonly EmailBodyResourceBudget _budget;
    private readonly EmailBodyResourceLimitState _limitState;
    private readonly CancellationToken _operationCancellationToken;
    private long _bytesRead;
    private bool _disposed;
    private bool _endOfStream;
    private bool _sourceFailed;

    internal EmailBodyResourceReadStream(
        Stream source,
        long maximumBytes,
        EmailBodyResourceBudget budget,
        EmailBodyResourceLimitState limitState,
        CancellationToken operationCancellationToken) {
        _source = source ?? throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("Source stream must be readable.", nameof(source));
        _maximumBytes = maximumBytes;
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
        _limitState = limitState ?? throw new ArgumentNullException(nameof(limitState));
        _operationCancellationToken = operationCancellationToken;
    }

    public override bool CanRead => !_disposed && _source.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) {
        ValidateReadArguments(buffer, offset, count);
        ThrowIfUnavailable();
        _operationCancellationToken.ThrowIfCancellationRequested();
        if (count == 0 || _endOfStream) return 0;
        return ReadCore(buffer, offset, count);
    }

    public override async Task<int> ReadAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        ValidateReadArguments(buffer, offset, count);
        ThrowIfUnavailable();
        _operationCancellationToken.ThrowIfCancellationRequested();
        cancellationToken.ThrowIfCancellationRequested();
        if (count == 0 || _endOfStream) return 0;
        CancellationTokenSource? linkedCancellation = null;
        CancellationToken effectiveCancellationToken;
        if (!_operationCancellationToken.CanBeCanceled ||
            _operationCancellationToken == cancellationToken) {
            effectiveCancellationToken = cancellationToken;
        } else if (!cancellationToken.CanBeCanceled) {
            effectiveCancellationToken = _operationCancellationToken;
        } else {
            linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                _operationCancellationToken,
                cancellationToken);
            effectiveCancellationToken = linkedCancellation.Token;
        }
        try {
            return await ReadCoreAsync(
                buffer,
                offset,
                count,
                effectiveCancellationToken).ConfigureAwait(false);
        } finally {
            linkedCancellation?.Dispose();
        }
    }

    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (!_disposed) {
            _disposed = true;
            if (disposing) _source.Dispose();
        }
        base.Dispose(disposing);
    }

    private int ReadCore(byte[] buffer, int offset, int count) {
        int allowed = GetAllowedReadCount(count);
        if (allowed == 0) return ProbeBoundary();
        int reserved = _budget.Reserve(allowed);
        if (reserved == 0) return ProbeAggregateBoundary();
        try {
            int read = _source.Read(buffer, offset, reserved);
            _budget.Commit(reserved, read);
            _bytesRead += read;
            if (read == 0) _endOfStream = true;
            return read;
        } catch {
            _budget.Release(reserved);
            _sourceFailed = true;
            throw;
        }
    }

    private async Task<int> ReadCoreAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        int allowed = GetAllowedReadCount(count);
        if (allowed == 0) return await ProbeBoundaryAsync(cancellationToken).ConfigureAwait(false);
        int reserved = _budget.Reserve(allowed);
        if (reserved == 0) return await ProbeAggregateBoundaryAsync(cancellationToken).ConfigureAwait(false);
        try {
            int read = await _source.ReadAsync(
                buffer,
                offset,
                reserved,
                cancellationToken).ConfigureAwait(false);
            _budget.Commit(reserved, read);
            _bytesRead += read;
            if (read == 0) _endOfStream = true;
            return read;
        } catch {
            _budget.Release(reserved);
            _sourceFailed = true;
            throw;
        }
    }

    private int ProbeBoundary() {
        int reserved = _budget.Reserve(1);
        if (reserved == 0) return ProbeAggregateBoundary();
        var probe = new byte[1];
        int read;
        try {
            read = _source.Read(probe, 0, 1);
        } catch {
            _budget.Release(1);
            _sourceFailed = true;
            throw;
        }
        _budget.Commit(1, read);
        if (read == 0) {
            _endOfStream = true;
            return 0;
        }
        return FailResourceLimit();
    }

    private async Task<int> ProbeBoundaryAsync(CancellationToken cancellationToken) {
        int reserved = _budget.Reserve(1);
        if (reserved == 0) return await ProbeAggregateBoundaryAsync(cancellationToken).ConfigureAwait(false);
        var probe = new byte[1];
        int read;
        try {
            read = await _source.ReadAsync(probe, 0, 1, cancellationToken).ConfigureAwait(false);
        } catch {
            _budget.Release(1);
            _sourceFailed = true;
            throw;
        }
        _budget.Commit(1, read);
        if (read == 0) {
            _endOfStream = true;
            return 0;
        }
        return FailResourceLimit();
    }

    private int ProbeAggregateBoundary() {
        var probe = new byte[1];
        try {
            int read = _source.Read(probe, 0, 1);
            if (read == 0) {
                _endOfStream = true;
                return 0;
            }
            return FailAggregateLimit();
        } catch {
            _sourceFailed = true;
            throw;
        }
    }

    private async Task<int> ProbeAggregateBoundaryAsync(CancellationToken cancellationToken) {
        var probe = new byte[1];
        try {
            int read = await _source.ReadAsync(probe, 0, 1, cancellationToken).ConfigureAwait(false);
            if (read == 0) {
                _endOfStream = true;
                return 0;
            }
            return FailAggregateLimit();
        } catch {
            _sourceFailed = true;
            throw;
        }
    }

    private int FailResourceLimit() {
        long actual = _maximumBytes == long.MaxValue ? long.MaxValue : _maximumBytes + 1L;
        _limitState.MarkExceeded(actual);
        _sourceFailed = true;
        throw new EmailLimitExceededException(
            "EmailBodyProjectionOptions.MaxResourceBytes",
            actual,
            _maximumBytes);
    }

    private int FailAggregateLimit() {
        _budget.MarkExceeded();
        _sourceFailed = true;
        _budget.ThrowIfExceeded();
        throw new InvalidOperationException("The aggregate resource limit failure was not recorded.");
    }

    private int GetAllowedReadCount(int requestedCount) {
        long remaining = _maximumBytes - _bytesRead;
        return remaining <= 0
            ? 0
            : (int)Math.Min(requestedCount, remaining);
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailBodyResourceReadStream));
        _limitState.ThrowIfExceeded();
        _budget.ThrowIfExceeded();
        if (_sourceFailed) {
            throw new InvalidOperationException("The resource stream cannot continue after a failed source read.");
        }
    }

    private static void ValidateReadArguments(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || offset > buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        if (buffer.Length - offset < count) throw new ArgumentException("Offset and count exceed the buffer length.");
    }
}
