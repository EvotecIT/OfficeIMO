using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>A durable provider reference. The bookmark grants access; the location identifies the document.</summary>
internal sealed record StudioStorageReference(string Location, string Name, string? Bookmark = null);

internal sealed record StudioStorageSnapshot(byte[] Bytes, string Identity);
internal sealed record StudioStoragePublication(string Fingerprint, string Identity);

/// <summary>Owns desktop provider items for the window lifetime and opens permission-scoped streams per operation.</summary>
internal sealed class StudioStorageAccess : IDisposable {
    internal const long MaximumDocumentBytes = 512L * 1024 * 1024;
    private readonly Dictionary<string, IStorageFile> _files = new(StringComparer.Ordinal);
    private readonly Dictionary<string, StudioStorageReference> _references = new(StringComparer.Ordinal);
    private readonly HashSet<IStorageFile> _retiredFiles = new(ReferenceEqualityComparer.Instance);
    private readonly object _sync = new();
    private Func<IStorageProvider>? _provider;
    private bool _disposed;

    internal void Attach(Func<IStorageProvider> provider) => _provider = provider ?? throw new ArgumentNullException(nameof(provider));

    internal async Task<string?> RegisterSingleAsync(IReadOnlyList<IStorageFile> files, CancellationToken token) {
        if (files.Count == 0) { token.ThrowIfCancellationRequested(); return null; }
        foreach (IStorageFile extra in files.Skip(1).Where(file => !ReferenceEquals(file, files[0]))
                     .Distinct<IStorageFile>(ReferenceEqualityComparer.Instance)) extra.Dispose();
        return await RegisterAsync(files[0], token).ConfigureAwait(true);
    }

    internal async Task<string> ReadIdentityAsync(string location, CancellationToken token) {
        await using Stream stream = await OpenReadAsync(location, token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        string? path = OfficeStorageIdentity.GetLocalPath(location);
        return path is null ? OfficeStorageIdentity.Normalize(location)
            : stream is FileStream file ? OfficePathIdentity.GetPhysicalIdentityKey(path, file.SafeFileHandle)
            : OfficePathIdentity.GetPhysicalIdentityKey(path);
    }

    internal async Task<string> RegisterAsync(IStorageFile file, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(file);
        bool retained = false;
        try {
            ObjectDisposedException.ThrowIf(_disposed, this);
            token.ThrowIfCancellationRequested();
            string location = Location(file);
            if (location.Length > 4096 || string.IsNullOrWhiteSpace(file.Name) || file.Name.Length > 4096) {
                throw new IOException("The provider's document reference exceeds the supported size.");
            }
            string key = OfficeStorageIdentity.Normalize(location);
            string? bookmark = file.CanBookmark ? await file.SaveBookmarkAsync().ConfigureAwait(true) : null;
            token.ThrowIfCancellationRequested();
            lock (_sync) {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (bookmark?.Length > 32768) bookmark = null;
                _references[key] = new(location, file.Name, bookmark);
                if (_files.TryGetValue(key, out IStorageFile? previous) && !ReferenceEquals(previous, file)) _retiredFiles.Add(previous);
                _files[key] = file;
                retained = true;
            }
            return location;
        } finally {
            if (!retained) file.Dispose();
        }
    }

    internal void Remember(StudioStorageReference reference) {
        if (reference.Bookmark?.Length > 32768 || reference.Name is null || reference.Name.Length > 4096) return;
        string location = OfficeStorageIdentity.Normalize(reference.Location);
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _references[OfficeStorageIdentity.Normalize(location)] = reference with { Location = location };
        }
    }

    internal StudioStorageReference Describe(string location) {
        lock (_sync) return _references.TryGetValue(OfficeStorageIdentity.Normalize(location), out var reference)
            ? reference : new(OfficeStorageIdentity.Normalize(location), OfficeStorageIdentity.GetFileName(location));
    }

    internal bool UsesProviderPublication(string location) {
        lock (_sync) return OfficeStorageIdentity.GetLocalPath(location) is null ||
            (OperatingSystem.IsMacOS() && _references.ContainsKey(OfficeStorageIdentity.Normalize(location)));
    }

    internal async Task<StudioStorageSnapshot> ReadSnapshotAsync(string location, CancellationToken token) {
        await using Stream stream = await OpenReadAsync(location, token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        string? localPath = OfficeStorageIdentity.GetLocalPath(location);
        string identity = localPath is null ? OfficeStorageIdentity.Normalize(location)
            : stream is FileStream file ? OfficePathIdentity.GetPhysicalIdentityKey(localPath, file.SafeFileHandle)
            : OfficePathIdentity.GetPhysicalIdentityKey(localPath);
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, token, MaximumDocumentBytes).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        if (localPath is not null && OfficePathIdentity.GetPhysicalIdentityKey(localPath) != identity) {
            throw new IOException("The document changed while it was being opened. Select it again to read the current file.");
        }
        return new(bytes, identity);
    }

    internal async Task<Stream> OpenReadAsync(string location, CancellationToken token) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        token.ThrowIfCancellationRequested();
        IStorageFile? file = await ResolveAsync(location, token).ConfigureAwait(false);
        if (file is not null) return await file.OpenReadAsync().ConfigureAwait(false);
        return new FileStream(OfficeStorageIdentity.GetLocalPath(location)!, FileMode.Open, FileAccess.Read,
            FileShare.Read | FileShare.Delete, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
    }

    internal Task<string> FingerprintAsync(string location, CancellationToken token) =>
        OfficeStreamPublication.ReadFingerprintAsync(ct => OpenReadAsync(location, ct), MaximumDocumentBytes, token);

    internal async Task<StudioStoragePublication> PublishAsync(string location, byte[] bytes, string? expectedFingerprint,
        Func<CancellationToken, Task> authorize, CancellationToken token) {
        string? identity = null;
        string fingerprint = await OfficeStreamPublication.WriteVerifiedAsync(async ct => {
            Stream stream = await OpenReadAsync(location, ct).ConfigureAwait(false);
            try {
                string? localPath = OfficeStorageIdentity.GetLocalPath(location);
                identity = localPath is null ? OfficeStorageIdentity.Normalize(location)
                    : stream is FileStream file ? OfficePathIdentity.GetPhysicalIdentityKey(localPath, file.SafeFileHandle)
                    : OfficePathIdentity.GetPhysicalIdentityKey(localPath);
                return stream;
            } catch { stream.Dispose(); throw; }
        }, async ct => {
            await authorize(ct).ConfigureAwait(false);
            IStorageFile file = await ResolveAsync(location, ct).ConfigureAwait(false)
                ?? throw new IOException("The storage provider is unavailable. Select the destination again.");
            ct.ThrowIfCancellationRequested();
            return await file.OpenWriteAsync().ConfigureAwait(false);
        }, bytes, expectedFingerprint, MaximumDocumentBytes, token).ConfigureAwait(false);
        return new(fingerprint, identity!);
    }

    private async Task<IStorageFile?> ResolveAsync(string location, CancellationToken token) {
        string key = OfficeStorageIdentity.Normalize(location);
        IStorageFile? existing;
        StudioStorageReference? reference;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_files.TryGetValue(key, out existing)) return existing;
            _references.TryGetValue(key, out reference);
        }
        if (reference?.Bookmark is null && OfficeStorageIdentity.GetLocalPath(location) is not null) return null;
        IStorageProvider provider = _provider?.Invoke()
            ?? throw new IOException("This document needs its storage provider. Select it again to grant access.");
        IStorageFile? file = reference?.Bookmark is { } bookmark
            ? await provider.OpenFileBookmarkAsync(bookmark).ConfigureAwait(false)
            : await provider.TryGetFileFromPathAsync(new Uri(location, UriKind.Absolute)).ConfigureAwait(false);
        if (file is null) throw new IOException("Access to this document is unavailable. Select it again to grant access.");
        bool retained = false;
        try {
            token.ThrowIfCancellationRequested();
            if (!string.Equals(OfficeStorageIdentity.Normalize(Location(file)), key, StringComparison.Ordinal)) {
                throw new IOException("The provider reference now identifies a different location. Select the document again.");
            }
            lock (_sync) {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_files.TryGetValue(key, out existing)) return existing;
                _files.Add(key, file);
                retained = true;
                return file;
            }
        } finally {
            if (!retained) file.Dispose();
        }
    }

    private static string Location(IStorageFile file) {
        if (file.TryGetLocalPath() is { } path) return OfficeStorageIdentity.Normalize(path);
        if (!file.Path.IsAbsoluteUri) throw new IOException("The provider did not supply a stable document location.");
        return OfficeStorageIdentity.Normalize(file.Path.AbsoluteUri);
    }

    public void Dispose() {
        IStorageFile[] files;
        lock (_sync) {
            if (_disposed) return;
            _disposed = true;
            files = _files.Values.Concat(_retiredFiles).Distinct<IStorageFile>(ReferenceEqualityComparer.Instance).ToArray();
            _files.Clear();
            _retiredFiles.Clear();
            _references.Clear();
        }
        foreach (IStorageFile file in files) file.Dispose();
    }
}
