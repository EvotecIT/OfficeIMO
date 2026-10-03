using System.Runtime.CompilerServices;
using Avalonia.Platform.Storage;
using OfficeIMO.Internal;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    private readonly Dictionary<string, IStorageFolder> _folders = new(StringComparer.Ordinal);
    private readonly HashSet<IStorageFolder> _retiredFolders = new(ReferenceEqualityComparer.Instance);

    internal bool IsFolder(string location) { lock (_sync) return _folders.ContainsKey(OfficeStorageIdentity.Normalize(location)); }

    internal Task<string?> RegisterFolderAsync(IReadOnlyList<IStorageFolder> folders, CancellationToken token) {
        if (folders.Count == 0) { token.ThrowIfCancellationRequested(); return Task.FromResult<string?>(null); }
        IStorageFolder selected = folders[0];
        foreach (var extra in folders.Skip(1).Distinct<IStorageFolder>(ReferenceEqualityComparer.Instance))
            if (!ReferenceEquals(extra, selected)) extra.Dispose();
        bool retained = false;
        try {
            token.ThrowIfCancellationRequested();
            string location = Location(selected);
            if (location.Length > 4096 || string.IsNullOrWhiteSpace(selected.Name) || selected.Name.Length > 4096)
                throw new IOException("The provider folder reference exceeds the supported size.");
            lock (_sync) {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_folders.TryGetValue(location, out var previous) && !ReferenceEquals(previous, selected)) _retiredFolders.Add(previous);
                _folders[location] = selected;
                _references[location] = new(location, selected.Name);
                retained = true;
            }
            return Task.FromResult<string?>(location);
        } finally { if (!retained) selected.Dispose(); }
    }

    internal DirectoryInputSession CreateDirectoryInputs(IEnumerable<string> locations) {
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return new(locations.Where(UsesProviderPublication).Where(IsFolder)
                .Distinct(StringComparer.Ordinal).ToDictionary(location => location, location => _folders[OfficeStorageIdentity.Normalize(location)], StringComparer.Ordinal));
        }
    }

    /// <summary>Keeps enumerated provider references alive for one operation, including publication verification.</summary>
    internal sealed class DirectoryInputSession : IDisposable {
        private readonly HashSet<IStorageItem> _items = new(ReferenceEqualityComparer.Instance);
        private readonly HashSet<IStorageItem> _roots;
        private bool _disposed;
        internal IReadOnlyDictionary<string, OfficeWorkflowDirectoryInput> Inputs { get; }
        private readonly IReadOnlyDictionary<string, IStorageFolder> _folders;

        internal DirectoryInputSession(IReadOnlyDictionary<string, IStorageFolder> folders) {
            _folders = folders;
            _roots = new(folders.Values, ReferenceEqualityComparer.Instance);
            Inputs = folders.ToDictionary(pair => pair.Key,
                pair => new OfficeWorkflowDirectoryInput((options, token) => Enumerate(pair.Value, options, token)), StringComparer.Ordinal);
        }

        internal OfficeWorkflowDirectoryPackageInput? CreatePackageInput(string location) =>
            _folders.TryGetValue(location, out var folder)
                ? new(folder.Name, Inputs[location], new PackageRootGuard(this, folder)) : null;

        private sealed class PackageRootGuard(DirectoryInputSession session, IStorageFolder folder) : IOfficeWorkflowPublicationGuard {
            private string? _identity;
            public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
                ObjectDisposedException.ThrowIf(session._disposed, session);
                // Avalonia keeps the folder's native permission scope open while its iterator is alive.
                // A reference to IStorageFolder alone does not grant filesystem access.
                await using var scope = folder.GetItemsAsync().GetAsyncEnumerator(token);
                if (!await scope.MoveNextAsync().ConfigureAwait(false))
                    throw new IOException("The selected package has no members.");
                if (!session._roots.Contains(scope.Current)) session._items.Add(scope.Current);
                token.ThrowIfCancellationRequested();
                string local = folder.TryGetLocalPath()
                    ?? throw new IOException("This provider cannot verify a directory package root. Select a local package.");
                if ((File.GetAttributes(local) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Linked directory packages cannot be included safely.");
                string identity = OfficePathIdentity.GetPhysicalIdentityKey(local);
                if (_identity is not null && _identity != identity)
                    throw new IOException("The source directory package was replaced during execution.");
                _identity ??= identity;
                if (OfficeStorageIdentity.AreEquivalent(Location(folder), path)) return false;
                string? output = OfficeStorageIdentity.GetLocalPath(path);
                if (output is null) return true;
                if (OfficePathIdentity.IsSameOrDescendant(output, local) ||
                    isDirectory && OfficePathIdentity.IsSameOrDescendant(local, output)) return false;
                // Batch protection also covers packages whose members have not yet been captured.
                if (!isDirectory && File.Exists(output)) {
                    await foreach (var entry in session.Enumerate(folder, new(true, 10_000), token).ConfigureAwait(false)) {
                        if (entry.Input is not null && OfficeStorageIdentity.AreEquivalent(entry.Location, path)) return false;
                    }
                }
                if (OfficePathIdentity.GetPhysicalIdentityKey(local) != _identity)
                    throw new IOException("The source directory package was replaced during execution.");
                return true;
            }
        }

        private async IAsyncEnumerable<OfficeWorkflowDirectoryEntry> Enumerate(IStorageFolder root,
            OfficeWorkflowDirectoryReadOptions options, [EnumeratorCancellation] CancellationToken token) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            int count = 0;
            var seen = new HashSet<string>(StringComparer.Ordinal) { Location(root) };
            await foreach (var entry in Walk(root, string.Empty, 0).WithCancellation(token).ConfigureAwait(false)) yield return entry;

            async IAsyncEnumerable<OfficeWorkflowDirectoryEntry> Walk(IStorageFolder folder, string prefix, int depth) {
                await foreach (var item in folder.GetItemsAsync().WithCancellation(token).ConfigureAwait(false)) {
                    if (!_roots.Contains(item)) _items.Add(item);
                    token.ThrowIfCancellationRequested();
                    if (++count > options.MaximumEntries) throw new IOException("The provider folder exceeds the workflow entry limit.");
                    if (depth >= options.MaximumDepth) throw new IOException("The provider folder exceeds the supported depth.");
                    string name = item.Name;
                    if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(['/', '\\', '\0']) >= 0 || name is "." or "..")
                        throw new IOException("The provider returned an invalid member name.");
                    string location = Location(item);
                    if (!seen.Add(location)) throw new IOException("The provider folder contains repeated or cyclic member identities.");
                    if (item.TryGetLocalPath() is { } local && (File.GetAttributes(local) & FileAttributes.ReparsePoint) != 0)
                        throw new IOException("Linked folder members cannot be included safely. Select the target files explicitly.");
                    string relative = prefix + name;
                    if (item is IStorageFile file) {
                        yield return new(relative, location, new(name, async cancellation => {
                            ObjectDisposedException.ThrowIf(_disposed, this);
                            cancellation.ThrowIfCancellationRequested();
                            return await file.OpenReadAsync().ConfigureAwait(false);
                        }));
                    } else if (item is IStorageFolder child) {
                        yield return new(relative, location, null);
                        if (options.IncludeSubdirectories)
                            await foreach (var entry in Walk(child, relative + "/", depth + 1).WithCancellation(token).ConfigureAwait(false)) yield return entry;
                    } else throw new IOException("The provider returned an unsupported folder member.");
                }
            }
        }

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            List<Exception>? errors = null;
            foreach (var item in _items) {
                try { item.Dispose(); }
                catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) { (errors ??= []).Add(error); }
            }
            _items.Clear();
            if (errors is not null) throw new IOException("Provider folder references could not be released.", new AggregateException(errors));
        }
    }
}
