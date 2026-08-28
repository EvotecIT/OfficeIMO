namespace OfficeIMO.Email;

/// <summary>Operation-scoped attachment resource resolved by CID, content location, or filename.</summary>
public sealed class EmailBodyResource {
    private readonly EmailAttachment _attachment;
    private readonly long _maximumBytes;
    private readonly EmailBodyResourceBudget _budget;

    internal EmailBodyResource(
        EmailAttachment attachment,
        long maximumBytes,
        EmailBodyResourceBudget budget) {
        _attachment = attachment ?? throw new ArgumentNullException(nameof(attachment));
        _maximumBytes = maximumBytes;
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
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
        cancellationToken.ThrowIfCancellationRequested();
        Stream source = _attachment.OpenContentStream();
        try {
            return new EmailBodyResourceReadStream(source, _maximumBytes, _budget, cancellationToken);
        } catch {
            source.Dispose();
            throw;
        }
    }

    /// <summary>Asynchronously opens a fresh sequential stream governed by the configured budgets.</summary>
    public async Task<Stream> OpenReadStreamAsync(CancellationToken cancellationToken = default) {
        EnsureDeclaredLengthAllowed();
        cancellationToken.ThrowIfCancellationRequested();
        Stream source = await _attachment.OpenContentStreamAsync(cancellationToken).ConfigureAwait(false);
        try {
            return new EmailBodyResourceReadStream(source, _maximumBytes, _budget, cancellationToken);
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

    internal EmailBodyResourceBudget(long maximumBytes) {
        _maximumBytes = maximumBytes;
    }

    internal void Consume(int count) {
        if (count <= 0) return;
        lock (_gate) {
            long next = checked(_consumedBytes + count);
            if (next > _maximumBytes) {
                throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxTotalResourceBytes",
                    next, _maximumBytes);
            }
            _consumedBytes = next;
        }
    }
}

internal sealed class EmailBodyResourceReadStream : Stream {
    private readonly Stream _source;
    private readonly long _maximumBytes;
    private readonly EmailBodyResourceBudget _budget;
    private readonly CancellationToken _operationCancellationToken;
    private long _bytesRead;
    private bool _disposed;

    internal EmailBodyResourceReadStream(
        Stream source,
        long maximumBytes,
        EmailBodyResourceBudget budget,
        CancellationToken operationCancellationToken) {
        _source = source ?? throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("Source stream must be readable.", nameof(source));
        _maximumBytes = maximumBytes;
        _budget = budget ?? throw new ArgumentNullException(nameof(budget));
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
        ThrowIfDisposed();
        _operationCancellationToken.ThrowIfCancellationRequested();
        int read = _source.Read(buffer, offset, count);
        Account(read);
        return read;
    }

    public override async Task<int> ReadAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        ThrowIfDisposed();
        _operationCancellationToken.ThrowIfCancellationRequested();
        cancellationToken.ThrowIfCancellationRequested();
        int read = await _source.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        Account(read);
        return read;
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

    private void Account(int count) {
        if (count <= 0) return;
        long next = checked(_bytesRead + count);
        if (next > _maximumBytes) {
            throw new EmailLimitExceededException("EmailBodyProjectionOptions.MaxResourceBytes",
                next, _maximumBytes);
        }
        _budget.Consume(count);
        _bytesRead = next;
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailBodyResourceReadStream));
    }
}
