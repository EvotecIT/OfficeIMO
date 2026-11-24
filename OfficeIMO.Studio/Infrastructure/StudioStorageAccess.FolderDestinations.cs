using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    /// <summary>Rejects a remembered folder destination that no longer lies inside its selected folder.</summary>
    private void EnsureFolderDestination(string location) {
        StudioStorageReference? reference;
        lock (_sync) _references.TryGetValue(OfficeStorageIdentity.Normalize(location), out reference);
        if (reference?.PermissionFolder is null) return;
        string root = OfficeStorageIdentity.GetLocalPath(reference.PermissionFolder)
            ?? throw new IOException("The saved folder permission is invalid.");
        string file = OfficeStorageIdentity.GetLocalPath(reference.Location)
            ?? throw new IOException("The saved folder destination is invalid.");
        if (!OfficePathIdentity.IsSameOrDescendant(file, root) || string.Equals(file, root, StringComparison.Ordinal))
            throw new IOException("The saved destination is outside its selected folder.");
    }

    internal string SelectLocalFileDestination(string folder, string name) {
        ValidateOutputName(name);
        string root = RequireLocalOutputFolder(folder);
        string location = OfficeStorageIdentity.Normalize(Path.Combine(root, name));
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_folders.ContainsKey(OfficeStorageIdentity.Normalize(folder)))
                throw new IOException("Select an output folder before choosing a file destination.");
            _references[location] = new(location, name, PermissionFolder: root);
        }
        return location;
    }
}
