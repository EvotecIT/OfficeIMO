using Avalonia.Platform.Storage;
using OfficeIMO.Internal;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    private readonly Dictionary<string, (IStorageFolder Folder, string Name)> _outputFolderReferences = new(StringComparer.Ordinal);

    private bool OwnsProviderItem(IStorageItem item) {
        lock (_sync) return _files.Values.Any(file => ReferenceEquals(file, item)) || _retiredFiles.Any(file => ReferenceEquals(file, item)) ||
            _folders.Values.Any(folder => ReferenceEquals(folder, item)) || _retiredFolders.Any(folder => ReferenceEquals(folder, item));
    }

    private void RememberFolderOutput(IStorageFolder folder, string name, string location) {
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            string key = OfficeStorageIdentity.Normalize(location);
            _references[key] = new(key, name);
            _outputFolderReferences[key] = (folder, name);
            if (_files.Remove(key, out var previous)) _retiredFiles.Add(previous);
        }
    }

    internal DirectoryOutputSession? CreateDirectoryOutput(string location, OfficeWorkflowOutputRecoveryStore recoveryStore) {
        if (!UsesProviderPublication(location)) return null;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_folders.TryGetValue(OfficeStorageIdentity.Normalize(location), out var folder))
                throw new IOException("Select the output folder again to grant provider access.");
            return new(this, folder, recoveryStore);
        }
    }

    /// <summary>Owns child references for one output operation; the selected folder remains window-owned.</summary>
    internal sealed class DirectoryOutputSession : IDisposable {
        private readonly IStorageFolder _folder;
        private readonly StudioStorageAccess _owner;
        private readonly OfficeWorkflowOutputRecoveryStore _recovery;
        private readonly HashSet<IStorageItem> _items = new(ReferenceEqualityComparer.Instance);
        private readonly HashSet<string> _names = new(StringComparer.OrdinalIgnoreCase);
        private bool _disposed;
        internal OfficeWorkflowDirectoryOutput Output { get; }

        internal DirectoryOutputSession(StudioStorageAccess owner, IStorageFolder folder, OfficeWorkflowOutputRecoveryStore recovery) {
            _owner = owner;
            _folder = folder;
            _recovery = recovery;
            Output = new(ResolveAsync);
        }

        internal async Task<OfficeWorkflowDirectoryOutputFile> ResolveAsync(string name, CancellationToken token) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name is "." or ".." ||
                name.IndexOfAny(['/', '\\', ':', '\0']) >= 0 || name.Any(char.IsControl))
                throw new IOException("The output filename is not a supported provider child name.");
            if (!_names.Add(name)) throw new IOException("Multiple outputs have the same filename. Choose distinct source names before publishing to this folder.");
            IStorageFile? initial = await FindAsync(name, token).ConfigureAwait(false);
            string? actualLocation = initial is null ? null : Location(initial);
            if (actualLocation is not null) _owner.RememberFolderOutput(_folder, name, actualLocation);
            string destination = actualLocation ?? Location(_folder);
            var output = new OfficeWorkflowStreamOutput(name, async cancellation => {
                if (actualLocation is null) throw new FileNotFoundException("The selected new output has not been created yet.");
                IStorageFile file = await FindAsync(name, cancellation).ConfigureAwait(false)
                    ?? throw new FileNotFoundException("The provider output is no longer available.");
                EnsureSameLocation(file, actualLocation);
                return await file.OpenReadAsync().ConfigureAwait(false);
            }, async cancellation => {
                if (actualLocation is null) throw new IOException("The provider output has not been prepared.");
                IStorageFile file = await FindAsync(name, cancellation).ConfigureAwait(false)
                    ?? throw new FileNotFoundException("The provider output is no longer available.");
                EnsureSameLocation(file, actualLocation);
                cancellation.ThrowIfCancellationRequested();
                return await file.OpenWriteAsync().ConfigureAwait(false);
            }, _recovery, initial is not null ? null : async cancellation => {
                // Creating a child may truncate it. The owner has already retained recovery before entering here.
                if (await FindAsync(name, cancellation).ConfigureAwait(false) is not null)
                    throw new IOException("An output with this name appeared after selection. Check the folder before trying again.");
                cancellation.ThrowIfCancellationRequested();
                IStorageFile created = await _folder.CreateFileAsync(name).ConfigureAwait(false)
                    ?? throw new IOException("The provider did not return the created output file.");
                Retain(created);
                cancellation.ThrowIfCancellationRequested();
                EnsureName(created, name);
                actualLocation = Location(created);
                _owner.RememberFolderOutput(_folder, name, actualLocation);
                return actualLocation;
            });
            return new(destination, output);
        }

        private async Task<IStorageFile?> FindAsync(string name, CancellationToken token) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            token.ThrowIfCancellationRequested();
            IStorageFile? file;
            try { file = await _folder.GetFileAsync(name).ConfigureAwait(false); }
            catch (FileNotFoundException) { return null; }
            if (file is not null) { Retain(file); EnsureName(file, name); }
            token.ThrowIfCancellationRequested();
            return file;
        }

        private void Retain(IStorageItem item) {
            if (!ReferenceEquals(item, _folder) && !_owner.OwnsProviderItem(item)) _items.Add(item);
        }
        private static void EnsureName(IStorageFile file, string name) {
            if (!string.Equals(file.Name, name, StringComparison.Ordinal)) throw new IOException("The provider returned a different output filename.");
        }
        private static void EnsureSameLocation(IStorageFile file, string location) {
            if (!string.Equals(Location(file), location, StringComparison.Ordinal)) throw new IOException("The provider output now identifies a different location.");
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
            if (errors is not null) throw new IOException("Provider output references could not be released.", new AggregateException(errors));
        }
    }
}
