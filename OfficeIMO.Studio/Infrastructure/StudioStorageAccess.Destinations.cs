using Avalonia.Platform.Storage;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    private readonly Dictionary<string, (IStorageFolder Folder, string Name)> _pendingDestinations = new(StringComparer.Ordinal);

    /// <summary>Resolves a selected filename without creating or truncating it before publication is authorized.</summary>
    internal async Task<string> SelectFileDestinationAsync(string folderLocation, string name, CancellationToken token) {
        ValidateOutputName(name);
        IStorageFolder folder;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            folder = _folders[OfficeStorageIdentity.Normalize(folderLocation)];
        }
        token.ThrowIfCancellationRequested();
        IStorageFile? existing;
        try { existing = await FindFolderFileAsync(folder, name, token).ConfigureAwait(false); }
        catch (FileNotFoundException) { existing = null; }
        if (existing is not null) {
            string selected = await RegisterAsync(existing, token).ConfigureAwait(false);
            lock (_sync) _pendingDestinations.Remove(selected);
            return selected;
        }
        string location = OfficeStorageIdentity.GetLocalPath(folderLocation) is { } local
            ? OfficeStorageIdentity.Normalize(Path.Combine(local, name))
            : OfficeStorageIdentity.Normalize(new Uri(new Uri(folderLocation.TrimEnd('/') + "/"), Uri.EscapeDataString(name)).AbsoluteUri);
        token.ThrowIfCancellationRequested();
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _pendingDestinations[location] = (folder, name);
            _references[location] = new(location, name);
        }
        return location;
    }

    internal override async Task<bool> FolderContainsAsync(string folderLocation, string name, CancellationToken token) {
        IStorageFile? existing;
        try { existing = await FindFolderFileAsync(SelectedFolder(folderLocation), name, token).ConfigureAwait(false); }
        catch (FileNotFoundException) { return false; }
        if (existing is null) return false;
        if (!OwnsProviderItem(existing)) existing.Dispose();
        return true;
    }

    internal override async Task<string> SelectNewFileAsync(string folderLocation, string name, CancellationToken token) {
        SelectedFolder(folderLocation);
        string location = await SelectFileDestinationAsync(folderLocation, name, token).ConfigureAwait(false);
        lock (_sync) if (_pendingDestinations.ContainsKey(location)) return location;
        throw new IOException("A file with this name appeared in the selected folder. Check the folder before trying again.");
    }

    private IStorageFolder SelectedFolder(string folderLocation) {
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return _folders.TryGetValue(OfficeStorageIdentity.Normalize(folderLocation), out var folder) ? folder
                : throw new IOException("Select the output folder again to grant provider access.");
        }
    }

    private async Task<IStorageFile> ResolveForWriteAsync(string location, CancellationToken token) {
        string key = OfficeStorageIdentity.Normalize(location);
        (IStorageFolder Folder, string Name) pending;
        lock (_sync) _pendingDestinations.TryGetValue(key, out pending);
        if (pending.Folder is null)
            return await ResolveAsync(location, token).ConfigureAwait(false)
                ?? throw new IOException("The storage provider is unavailable. Select the destination again.");
        token.ThrowIfCancellationRequested();
        IStorageFile? appeared;
        try { appeared = await FindFolderFileAsync(pending.Folder, pending.Name, token).ConfigureAwait(false); }
        catch (FileNotFoundException) { appeared = null; }
        if (appeared is not null) {
            if (!OwnsProviderItem(appeared)) appeared.Dispose();
            throw new IOException("A file with this name appeared after selection. Select the destination again before replacing it.");
        }
        token.ThrowIfCancellationRequested();
        IStorageFile created = await pending.Folder.CreateFileAsync(pending.Name).ConfigureAwait(false)
            ?? throw new IOException("The provider did not return the created output file.");
        if (Location(created) != key || created.Name != pending.Name) {
            created.Dispose();
            throw new IOException("The provider created a different output location. Check the folder before trying again.");
        }
        await RegisterAsync(created, token).ConfigureAwait(false);
        lock (_sync) _pendingDestinations.Remove(key);
        return created;
    }
}
