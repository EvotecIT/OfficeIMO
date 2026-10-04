using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    private EmailReadWorkspace? _snapshotWorkspace;
    private string? _snapshotFingerprint;

    /// <summary>True when this session owns a private read-only copy of the source bytes.</summary>
    public bool IsSnapshot { get { ThrowIfDisposed(); return _snapshotWorkspace != null; } }

    /// <summary>
    /// Copies one store file into session-owned temporary storage and hashes it while copying.
    /// Repeated durable queries reuse that complete-source hash. The copy is deleted on disposal.
    /// Directory stores use their existing catalog and content validation instead.
    /// </summary>
    public static EmailStoreSession OpenSnapshot(string path, EmailStoreReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            64 * 1024, FileOptions.SequentialScan);
        return OpenSnapshot(source, Path.GetFileName(path), options, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Copies a readable stream into private temporary storage. Seekable sources are copied from their beginning
    /// and restored; non-seekable sources are copied forward. The snapshot represents the bytes observed while
    /// copying; callers must keep the original source stable until this call returns.
    /// </summary>
    public static EmailStoreSession OpenSnapshot(Stream stream, string? sourceName = null,
        EmailStoreReaderOptions? options = null, bool leaveOpen = true,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source must be readable.", nameof(stream));
        EmailStoreReaderOptions effective = options ?? EmailStoreReaderOptions.Default;
        long? originalPosition = stream.CanSeek ? stream.Position : (long?)null;
        EmailReadWorkspace? workspace = null;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (stream.CanSeek) {
                if (stream.Length > effective.MaxInputBytes)
                    throw new EmailStoreLimitExceededException(nameof(effective.MaxInputBytes), stream.Length, effective.MaxInputBytes);
                stream.Position = 0;
            }
            workspace = new EmailReadWorkspace();
            string snapshotPath = workspace.CreateInputPath();
            string fingerprint;
            using (var output = new FileStream(snapshotPath, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                       64 * 1024, FileOptions.SequentialScan))
            using (IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256)) {
                var buffer = new byte[64 * 1024];
                long total = 0;
                while (true) {
                    cancellationToken.ThrowIfCancellationRequested();
                    long remaining = effective.MaxInputBytes - total;
                    int count = stream.Read(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1);
                    if (count == 0) break;
                    total += count;
                    if (total > effective.MaxInputBytes)
                        throw new EmailStoreLimitExceededException(nameof(effective.MaxInputBytes), total, effective.MaxInputBytes);
                    output.Write(buffer, 0, count);
                    hash.AppendData(buffer, 0, count);
                }
                cancellationToken.ThrowIfCancellationRequested();
                fingerprint = EmailHashing.ToHexLower(hash.GetHashAndReset());
            }
            var snapshot = new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.None,
                64 * 1024, FileOptions.RandomAccess);
            EmailStoreSession session;
            try {
                session = OpenCore(snapshot, sourceName, effective,
                    leaveOpen: false, originalPosition: 0, cancellationToken, isSnapshot: true);
            } catch {
                snapshot.Dispose();
                throw;
            }
            session._snapshotWorkspace = workspace;
            session._snapshotFingerprint = fingerprint;
            workspace = null;
            return session;
        } finally {
            workspace?.Dispose();
            if (leaveOpen) {
                if (originalPosition.HasValue) stream.Position = originalPosition.Value;
            } else {
                stream.Dispose();
            }
        }
    }
}
