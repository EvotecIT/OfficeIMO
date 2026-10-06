using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>A durable provider reference. The bookmark grants access; the location identifies the document.</summary>
internal sealed record StudioStorageReference(string Location, string Name, string? Bookmark = null);

internal sealed record StudioStorageSnapshot(byte[] Bytes, string Identity);
internal sealed record StudioStoragePublication(string Fingerprint, string Identity);

/// <summary>Shared document reads, identity checks and verified publication over permission-scoped streams.</summary>
internal class StudioDocumentStorage : IDisposable {
    internal const long MaximumDocumentBytes = 512L * 1024 * 1024;
    private readonly string? _protectedRecoveryRoot;
    internal StudioDocumentStorage(string? protectedRecoveryRoot = null) => _protectedRecoveryRoot = protectedRecoveryRoot;
    internal bool IsRecoveryLocation(string location, bool isDirectory = false) {
        string? path = OfficeStorageIdentity.GetLocalPath(location);
        return path is not null && _protectedRecoveryRoot is not null &&
            (OfficePathIdentity.IsSameOrDescendant(path, _protectedRecoveryRoot) ||
             isDirectory && OfficePathIdentity.IsSameOrDescendant(_protectedRecoveryRoot, path));
    }

    internal void EnsureWritableLocation(string location) {
        if (IsRecoveryLocation(location)) throw new IOException("Recovery copies are protected. Use Save As to save your changes to another location.");
    }

    internal virtual void Remember(StudioStorageReference reference) { }
    internal virtual StudioStorageReference Describe(string location) => new(OfficeStorageIdentity.Normalize(location), OfficeStorageIdentity.GetFileName(location));
    internal virtual bool UsesProviderPublication(string location) => OfficeStorageIdentity.GetLocalPath(location) is null;
    internal virtual Task<Stream> OpenReadAsync(string location, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        return Task.FromResult<Stream>(new FileStream(OfficeStorageIdentity.GetLocalPath(location)
            ?? throw new IOException("This document needs its storage provider."), FileMode.Open, FileAccess.Read,
            FileShare.Read | FileShare.Delete, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan));
    }
    internal virtual Task<Stream> OpenWriteAsync(string location, CancellationToken token) =>
        throw new IOException("The storage provider is unavailable. Select the destination again.");
    internal virtual string OpenedIdentity(string location, Stream stream) {
        string? path = OfficeStorageIdentity.GetLocalPath(location);
        if (path is null) return OfficeStorageIdentity.Normalize(location);
        string identity = stream is FileStream file
            ? OfficePathIdentity.GetPhysicalIdentityKey(path, file.SafeFileHandle)
            : OfficePathIdentity.GetPhysicalIdentityKey(path);
        if (OfficePathIdentity.GetPhysicalIdentityKey(path) != identity)
            throw new IOException("The document changed while it was being opened. Select it again to read the current file.");
        return identity;
    }
    internal async Task<string> ReadIdentityAsync(string location, CancellationToken token) {
        await using Stream stream = await OpenReadAsync(location, token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        return OpenedIdentity(location, stream);
    }

    internal async Task<StudioStorageSnapshot> ReadSnapshotAsync(string location, CancellationToken token,
        long maximumBytes = MaximumDocumentBytes) {
        if (maximumBytes < 1 || maximumBytes > MaximumDocumentBytes) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        await using Stream stream = await OpenReadAsync(location, token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        string? localPath = OfficeStorageIdentity.GetLocalPath(location);
        string identity = OpenedIdentity(location, stream);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, token, maximumBytes).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        if (localPath is not null && OfficePathIdentity.GetPhysicalIdentityKey(localPath) != identity) {
            throw new IOException("The document changed while it was being opened. Select it again to read the current file.");
        }
        return new(bytes, identity);
    }

    internal Task<string> FingerprintAsync(string location, CancellationToken token) =>
        OfficeStreamPublication.ReadFingerprintAsync(ct => OpenReadAsync(location, ct), MaximumDocumentBytes, token);

    internal async Task<StudioStoragePublication> PublishAsync(string location, byte[] bytes, string? expectedFingerprint,
        Func<CancellationToken, Task> authorize, CancellationToken token) {
        EnsureWritableLocation(location);
        string? identity = null;
        string fingerprint = await OfficeStreamPublication.WriteVerifiedAsync(async ct => {
            Stream stream = await OpenReadAsync(location, ct).ConfigureAwait(false);
            try {
                identity = OpenedIdentity(location, stream);
                return stream;
            } catch { stream.Dispose(); throw; }
        }, async ct => {
            await authorize(ct).ConfigureAwait(false);
            return await OpenWriteAsync(location, ct).ConfigureAwait(false);
        }, bytes, expectedFingerprint, MaximumDocumentBytes, token).ConfigureAwait(false);
        return new(fingerprint, identity!);
    }

    internal static void ValidateOutputName(string name) {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name is "." or ".." ||
            name.IndexOfAny(['/', '\\', ':', '\0']) >= 0 || name.Any(char.IsControl))
            throw new IOException("Choose a filename without slashes, colons or control characters (up to 255 characters).");
    }

    public virtual void Dispose() { }
}
