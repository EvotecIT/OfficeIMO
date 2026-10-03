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

    internal string? Resolve(string identity) {
        if (!identity.StartsWith(Prefix, StringComparison.Ordinal)) return null;
        string relative = Uri.UnescapeDataString(identity[Prefix.Length..]);
        if (System.IO.Path.IsPathFullyQualified(relative)) return null;
        string full = System.IO.Path.GetFullPath(System.IO.Path.Combine(Path, relative));
        return full.StartsWith(Path + System.IO.Path.DirectorySeparatorChar, StringComparison.Ordinal) ? full : null;
    }
}
