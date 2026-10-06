using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Owns desktop provider items for the window lifetime and opens permission-scoped streams per operation.</summary>
internal sealed partial class StudioStorageAccess : StudioDocumentStorage {
    private readonly Dictionary<string, IStorageFile> _files = new(StringComparer.Ordinal);
    private readonly Dictionary<string, StudioStorageReference> _references = new(StringComparer.Ordinal);
    private readonly HashSet<IStorageFile> _retiredFiles = new(ReferenceEqualityComparer.Instance);
    private readonly object _sync = new();
    private Func<IStorageProvider>? _provider;
    private bool _disposed;

    internal StudioStorageAccess(string? protectedRecoveryRoot = null) : base(protectedRecoveryRoot) { }

    internal void Attach(Func<IStorageProvider> provider) => _provider = provider ?? throw new ArgumentNullException(nameof(provider));

    internal async Task<string?> RegisterSingleAsync(IReadOnlyList<IStorageFile> files, CancellationToken token) {
        if (files.Count == 0) { token.ThrowIfCancellationRequested(); return null; }
        foreach (IStorageFile extra in files.Skip(1).Where(file => !ReferenceEquals(file, files[0]))
                     .Distinct<IStorageFile>(ReferenceEqualityComparer.Instance)) extra.Dispose();
        return await RegisterAsync(files[0], token).ConfigureAwait(true);
    }

    internal async Task<IReadOnlyList<string>> RegisterManyAsync(IReadOnlyList<IStorageItem> files, CancellationToken token) {
        IStorageItem[] distinct = files.Distinct<IStorageItem>(ReferenceEqualityComparer.Instance).ToArray();
        var locations = new List<string>(distinct.Length);
        int index = 0;
        try {
            for (; index < distinct.Length; index++) {
                string? location;
                if (distinct[index] is IStorageFile file) location = await RegisterAsync(file, token).ConfigureAwait(true);
                else if (distinct[index] is IStorageFolder folder) location = await RegisterFolderAsync([folder], token).ConfigureAwait(true);
                else { distinct[index].Dispose(); throw new IOException("The provider returned an unsupported document item."); }
                if (location is not null) locations.Add(location);
            }
            token.ThrowIfCancellationRequested();
            return locations;
        } finally {
            // RegisterAsync owns the current item even on failure; release items it never reached.
            for (int remaining = index + 1; remaining < distinct.Length; remaining++) distinct[remaining].Dispose();
        }
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
            if (OperatingSystem.IsMacOS() && OfficeMacFilePermission.IsSandboxed &&
                OfficeStorageIdentity.GetLocalPath(location) is { } localPath && File.Exists(localPath)) {
                await using Stream access = await file.OpenReadAsync().ConfigureAwait(true);
                bookmark = OfficeMacFilePermission.CreateBookmark(localPath);
            }
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

    internal override void Remember(StudioStorageReference reference) {
        if (reference.Bookmark?.Length > 32768 || reference.Name is null || reference.Name.Length > 4096) return;
        string location = OfficeStorageIdentity.Normalize(reference.Location);
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _references[OfficeStorageIdentity.Normalize(location)] = reference with { Location = location };
        }
    }

    internal override StudioStorageReference Describe(string location) {
        lock (_sync) return _references.TryGetValue(OfficeStorageIdentity.Normalize(location), out var reference)
            ? reference : new(OfficeStorageIdentity.Normalize(location), OfficeStorageIdentity.GetFileName(location));
    }

    internal override bool UsesProviderPublication(string location) {
        string key = OfficeStorageIdentity.Normalize(location);
        lock (_sync) return _pendingDestinations.ContainsKey(key) || OfficeStorageIdentity.GetLocalPath(location) is null ||
            ((OperatingSystem.IsMacOS() || OperatingSystem.IsIOS()) && (_files.ContainsKey(key) || _folders.ContainsKey(key) ||
                _outputFolderReferences.ContainsKey(key) ||
                (_references.TryGetValue(key, out var reference) && reference.Bookmark is not null)));
    }

    internal OfficeIMO.Workflows.OfficeWorkflowStreamInput? CreateWorkflowInput(string location) =>
        UsesProviderPublication(location)
            ? new(Describe(location).Name, token => OpenReadAsync(location, token)) : null;

    internal OfficeIMO.Workflows.OfficeWorkflowStreamOutput? CreateWorkflowOutput(string location,
        OfficeIMO.Workflows.OfficeWorkflowOutputRecoveryStore recoveryStore) =>
        UsesProviderPublication(location)
            ? new(Describe(location).Name, token => OpenReadAsync(location, token),
                token => OpenWriteAsync(location, token), recoveryStore) : null;

    internal override async Task<Stream> OpenReadAsync(string location, CancellationToken token) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        token.ThrowIfCancellationRequested();
        OfficeMacFilePermission? permission = OpenNativePermission(location);
        try {
            IStorageFile? file = await ResolveAsync(location, token).ConfigureAwait(false);
            Stream stream = file is not null ? await file.OpenReadAsync().ConfigureAwait(false)
                : new FileStream(OfficeStorageIdentity.GetLocalPath(location)!, FileMode.Open, FileAccess.Read,
                    FileShare.Read | FileShare.Delete, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
            try {
                token.ThrowIfCancellationRequested();
                if (permission is null) RefreshNativePermission(location);
            }
            catch { await stream.DisposeAsync(); throw; }
            return permission is null ? stream : new PermissionStream(stream, permission);
        } catch { permission?.Dispose(); throw; }
    }

    private async Task<IStorageFile?> ResolveAsync(string location, CancellationToken token) {
        string key = OfficeStorageIdentity.Normalize(location);
        IStorageFile? existing;
        StudioStorageReference? reference;
        (IStorageFolder Folder, string Name) folderOutput;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_files.TryGetValue(key, out existing)) return existing;
            if (_pendingDestinations.ContainsKey(key)) throw new FileNotFoundException("The selected new output has not been created yet.");
            _references.TryGetValue(key, out reference);
            _outputFolderReferences.TryGetValue(key, out folderOutput);
        }
        if (folderOutput.Folder is null && reference?.Bookmark is null && OfficeStorageIdentity.GetLocalPath(location) is not null) return null;
        IStorageProvider? provider = folderOutput.Folder is not null ? null : _provider?.Invoke()
            ?? throw new IOException("This document needs its storage provider. Select it again to grant access.");
        IStorageFile? file = folderOutput.Folder is not null ? await FindFolderFileAsync(folderOutput.Folder, folderOutput.Name, token).ConfigureAwait(false)
            : reference?.Bookmark is { } native && OfficeMacFilePermission.IsNativeBookmark(native)
            ? await provider!.TryGetFileFromPathAsync(new Uri(location, UriKind.Absolute)).ConfigureAwait(false)
            : reference?.Bookmark is { } bookmark
            ? await provider!.OpenFileBookmarkAsync(bookmark).ConfigureAwait(false)
            : await provider!.TryGetFileFromPathAsync(new Uri(location, UriKind.Absolute)).ConfigureAwait(false);
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

    private static string Location(IStorageItem file) {
        if (file.TryGetLocalPath() is { } path) return OfficeStorageIdentity.Normalize(path);
        if (!file.Path.IsAbsoluteUri) throw new IOException("The provider did not supply a stable document location.");
        return OfficeStorageIdentity.Normalize(file.Path.AbsoluteUri);
    }

    public override void Dispose() {
        IStorageItem[] files;
        lock (_sync) {
            if (_disposed) return;
            _disposed = true;
            files = _files.Values.Concat(_retiredFiles).Cast<IStorageItem>()
                .Concat(_folders.Values).Concat(_retiredFolders).Distinct<IStorageItem>(ReferenceEqualityComparer.Instance).ToArray();
            _files.Clear();
            _retiredFiles.Clear();
            _references.Clear();
            _folders.Clear();
            _outputFolderReferences.Clear();
            _pendingDestinations.Clear();
            _retiredFolders.Clear();
        }
        List<Exception>? errors = null;
        foreach (IStorageItem file in files) {
            try { file.Dispose(); }
            catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) { (errors ??= []).Add(error); }
        }
        if (errors is not null) throw new IOException("Provider references could not be released.", new AggregateException(errors));
    }
}
