using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    // The window owns the selected folder and its native access together. Raw filesystem
    // writers, including staging and atomic rename, retain that access until the window closes.
    private readonly Dictionary<string, OfficeMacFilePermission> _folderPermissions = new(StringComparer.Ordinal);

    private static bool NeedsFolderPermission => OperatingSystem.IsIOS() || OfficeMacFilePermission.IsSandboxed;

    private static OfficeMacFilePermission? OpenFolderPermission(string folder, string? bookmark) =>
        NeedsFolderPermission
            ? OfficeMacFilePermission.OpenBookmarkData(DecodeFolderBookmark(bookmark, OperatingSystem.IsIOS() ? "ios" : "macOS"), folder)
            : null;

    /// <summary>Reads Avalonia's versioned picker envelope; the native owner verifies the resolved path.</summary>
    internal static byte[] DecodeFolderBookmark(string? bookmark, string platform) {
        if (bookmark is null || bookmark.Length > 32768 || platform is not ("ios" or "macOS"))
            throw new IOException("The folder permission is unavailable. Select the folder again.");
        byte[] bytes;
        try { bytes = Convert.FromBase64String(bookmark); }
        catch (FormatException error) { throw new IOException("The folder permission is invalid. Select the folder again.", error); }
        byte[] header = new byte[16];
        Encoding.ASCII.GetBytes("ava.v1." + platform).CopyTo(header, 0);
        if (bytes.Length <= header.Length || bytes.Length > header.Length + 24000 || !bytes.AsSpan(0, header.Length).SequenceEqual(header))
            throw new IOException("The folder permission is invalid or belongs to another platform. Select the folder again.");
        return bytes.AsSpan(header.Length).ToArray();
    }

    private void RetainFolderPermission(string folder, OfficeMacFilePermission? permission) {
        if (permission is null) return;
        if (_folderPermissions.Remove(folder, out var previous)) previous.Dispose();
        _folderPermissions.Add(folder, permission);
    }

    private void EnsureFolderPermission(StudioStorageReference reference) {
        string root = OfficeStorageIdentity.GetLocalPath(reference.PermissionFolder!)
            ?? throw new IOException("The saved folder permission is invalid.");
        string file = OfficeStorageIdentity.GetLocalPath(reference.Location)
            ?? throw new IOException("The saved folder destination is invalid.");
        OfficeMacFilePermission? permission = null;
        try {
            lock (_sync) {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (!_folderPermissions.ContainsKey(root)) permission = OpenFolderPermission(root, reference.Bookmark);
                if (!OfficePathIdentity.IsSameOrDescendant(file, root) || string.Equals(file, root, StringComparison.Ordinal))
                    throw new IOException("The saved destination is outside its selected folder.");
                RetainFolderPermission(root, permission);
                permission = null;
            }
        } finally { permission?.Dispose(); }
    }

    internal string SelectLocalFileDestination(string folder, string name) {
        ValidateOutputName(name);
        string root = RequireLocalOutputFolder(folder);
        string location = OfficeStorageIdentity.Normalize(Path.Combine(root, name));
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_folders.ContainsKey(OfficeStorageIdentity.Normalize(folder)))
                throw new IOException("Select an output folder before choosing a file destination.");
            var selected = _references[OfficeStorageIdentity.Normalize(folder)];
            _references[location] = new(location, name, selected.Bookmark, root);
        }
        return location;
    }
}
