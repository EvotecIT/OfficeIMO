using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO;

/// <summary>One bounded inert payload inventory and cumulative read operation.</summary>
/// <remarks>Keep the native document alive while requesting new reads. Each returned read-only stream is a complete detached copy owned by the caller and survives disposal of the native document and this session. Payloads are never executed and external targets are never resolved.</remarks>
public sealed class OfficeEmbeddedPayloadInspection : IDisposable {
    private readonly object _gate = new object();
    private readonly OfficeEmbeddedPayloadInspectionOptions _options;
    private readonly OfficeEmbeddedPayloadSource[] _sources;
    private readonly long?[] _observedLengths;
    private long _totalPayloadBytes;
    private long _readBytes;
    private bool _disposed;

    internal OfficeEmbeddedPayloadInspection(IEnumerable<OfficeEmbeddedPayloadSource> sources,
        OfficeEmbeddedPayloadInspectionOptions? options, CancellationToken cancellationToken) {
        _options = (options ?? new OfficeEmbeddedPayloadInspectionOptions()).Snapshot();
        var admitted = new List<OfficeEmbeddedPayloadSource>();
        foreach (OfficeEmbeddedPayloadSource source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (admitted.Count >= _options.MaxPayloads) throw Limit(nameof(_options.MaxPayloads), _options.MaxPayloads);
            admitted.Add(source);
        }
        _sources = admitted.ToArray();
        _observedLengths = new long?[_sources.Length];
        var descriptors = new OfficeEmbeddedPayloadInfo[_sources.Length];
        for (int ordinal = 0; ordinal < _sources.Length; ordinal++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeEmbeddedPayloadInfo info = _sources[ordinal].Info;
            if (info.HasKnownLength) AdmitLength(ordinal, info.Length);
            string? hash = null;
            if (_options.IncludeSha256) {
                using SHA256 sha = SHA256.Create();
                ReadPayload(ordinal, output: null, sha, cancellationToken);
                hash = string.Concat(sha.Hash!.Select(value => value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
            }
            descriptors[ordinal] = CopyInfo(info, ordinal, _observedLengths[ordinal] ?? -1, hash);
        }
        Payloads = Array.AsReadOnly(descriptors);
    }

    /// <summary>Ordered native descriptors. Ordinals address this session only; identities remain native owner/relationship or source-profile identities.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> Payloads { get; }
    /// <summary>Actual decoded bytes consumed by completed or failed reads and requested hashes.</summary>
    public long ReadBytes { get { lock (_gate) return _readBytes; } }

    /// <summary>Reads one complete payload within the snapshotted member and cumulative operation ceilings.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A member, aggregate payload, or cumulative read ceiling is reached.</exception>
    public Stream OpenRead(int ordinal, CancellationToken cancellationToken = default) {
        lock (_gate) {
            ThrowIfDisposed();
            if (ordinal < 0 || ordinal >= _sources.Length) throw new ArgumentOutOfRangeException(nameof(ordinal));
            cancellationToken.ThrowIfCancellationRequested();
            if (_readBytes > _options.MaxReadBytes) throw Limit(nameof(_options.MaxReadBytes), _options.MaxReadBytes);
            using var output = new MemoryStream();
            ReadPayload(ordinal, output, hash: null, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return new MemoryStream(output.GetBuffer(), 0, checked((int)output.Length), writable: false, publiclyVisible: false);
        }
    }

    /// <summary>Ends new-read access. Already-returned detached streams remain usable.</summary>
    public void Dispose() {
        lock (_gate) {
            if (_disposed) return;
            _disposed = true;
            Array.Clear(_sources, 0, _sources.Length);
        }
    }

    private void ReadPayload(int ordinal, Stream? output, HashAlgorithm? hash, CancellationToken token) {
        _sources[ordinal].ExecuteRead(() => ReadPayloadCore(ordinal, output, hash, token));
    }

    private void ReadPayloadCore(int ordinal, Stream? output, HashAlgorithm? hash, CancellationToken token) {
        using Stream source = _sources[ordinal].OpenRead(token);
        if (source == null || !source.CanRead) throw new InvalidDataException("The native payload source is not readable.");
        var buffer = new byte[81920];
        long length = 0;
        while (true) {
            token.ThrowIfCancellationRequested();
            long remainingMember = _options.MaxPayloadBytes - length;
            long remainingWork = _options.MaxReadBytes - _readBytes;
            long remainingTotal = _options.MaxTotalPayloadBytes - (_totalPayloadBytes - (_observedLengths[ordinal] ?? 0)) - length;
            long remaining = Math.Min(remainingMember, Math.Min(remainingWork, remainingTotal));
            int request = remaining >= buffer.Length ? buffer.Length : checked((int)Math.Max(1, remaining + 1));
            int read = source.Read(buffer, 0, request);
            if (read == 0) break;
            _readBytes = checked(_readBytes + read);
            length = checked(length + read);
            if (output != null && length > int.MaxValue) throw Limit("MaximumDetachedPayloadBytes", int.MaxValue);
            if (length > _options.MaxPayloadBytes) throw Limit(nameof(_options.MaxPayloadBytes), _options.MaxPayloadBytes);
            if (_readBytes > _options.MaxReadBytes) throw Limit(nameof(_options.MaxReadBytes), _options.MaxReadBytes);
            if (length > _options.MaxTotalPayloadBytes - (_totalPayloadBytes - (_observedLengths[ordinal] ?? 0)))
                throw Limit(nameof(_options.MaxTotalPayloadBytes), _options.MaxTotalPayloadBytes);
            output?.Write(buffer, 0, read);
            hash?.TransformBlock(buffer, 0, read, buffer, 0);
        }
        token.ThrowIfCancellationRequested();
        if (_sources[ordinal].Info.HasKnownLength && length != _sources[ordinal].Info.Length)
            throw new InvalidDataException("Native payload length differs from its declared metadata.");
        AdmitLength(ordinal, length);
        hash?.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
    }

    private void AdmitLength(int ordinal, long length) {
        if (length > _options.MaxPayloadBytes) throw Limit(nameof(_options.MaxPayloadBytes), _options.MaxPayloadBytes);
        long prior = _observedLengths[ordinal] ?? 0;
        long other = _totalPayloadBytes - prior;
        if (length > _options.MaxTotalPayloadBytes - other) throw Limit(nameof(_options.MaxTotalPayloadBytes), _options.MaxTotalPayloadBytes);
        _totalPayloadBytes = other + length;
        _observedLengths[ordinal] = length;
    }

    private static OfficeEmbeddedPayloadInfo CopyInfo(OfficeEmbeddedPayloadInfo info, int ordinal, long length, string? sha256) =>
        new OfficeEmbeddedPayloadInfo(info.Id, info.Kind, info.OwnerPartUri, info.RelationshipId, info.PartUri,
            info.ContentType, info.SuggestedFileName, length, sha256, ordinal);
    private static OfficeInspectionResourceLimitException Limit(string name, long maximum) => new OfficeInspectionResourceLimitException("embedded-payload", name, maximum);
    private void ThrowIfDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(OfficeEmbeddedPayloadInspection)); }
}
