namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Gives app-owned documents stable identities when the operating system relocates their container.</summary>
internal sealed class StudioLocalDocumentRoot(string root) {
    private const string Prefix = "officeimo-local://documents/";
    internal string Path { get; } = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar);

    internal string GetIdentity(string path) {
        if (!System.IO.Path.IsPathFullyQualified(path)) return path;
        string full = System.IO.Path.GetFullPath(path);
        return full.StartsWith(Path + System.IO.Path.DirectorySeparatorChar, StringComparison.Ordinal)
            ? Prefix + Uri.EscapeDataString(System.IO.Path.GetRelativePath(Path, full))
            : full;
    }

    /// <summary>Stores an imported document as a separate app-owned working copy.</summary>
    internal async Task<string> WriteWorkingCopyAsync(string name, byte[] bytes, CancellationToken token) {
        StudioDocumentStorage.ValidateOutputName(name);
        string directory = System.IO.Path.Combine(Path, Guid.NewGuid().ToString("N"));
        string destination = System.IO.Path.Combine(directory, name);
        await Task.Run(() => {
            token.ThrowIfCancellationRequested();
            Directory.CreateDirectory(directory);
            try {
                OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(destination, bytes,
                    OfficeIMO.Core.Internal.OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
                token.ThrowIfCancellationRequested();
            } catch {
                if (File.Exists(destination)) File.Delete(destination);
                Directory.Delete(directory);
                throw;
            }
        }, token);
        return destination;
    }

    /// <summary>Removes an unaccepted fresh import from its isolated app-owned directory.</summary>
    internal void DiscardWorkingCopy(string path) {
        string directory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path))!;
        if (System.IO.Path.GetDirectoryName(directory) != Path ||
            !Guid.TryParseExact(System.IO.Path.GetFileName(directory), "N", out _))
            throw new IOException("The discarded import is not in an app-owned working-copy directory.");
        File.Delete(path);
        Directory.Delete(directory);
    }

    internal static bool IsIdentity(string identity) => identity.StartsWith(Prefix, StringComparison.Ordinal);

    internal string? Resolve(string identity) {
        if (!identity.StartsWith(Prefix, StringComparison.Ordinal)) return null;
        string relative = Uri.UnescapeDataString(identity[Prefix.Length..]);
        if (System.IO.Path.IsPathFullyQualified(relative)) return null;
        string full = System.IO.Path.GetFullPath(System.IO.Path.Combine(Path, relative));
        return full.StartsWith(Path + System.IO.Path.DirectorySeparatorChar, StringComparison.Ordinal) ? full : null;
    }
}
