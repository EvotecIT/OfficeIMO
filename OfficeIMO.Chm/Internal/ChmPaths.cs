namespace OfficeIMO.Chm;

internal static class ChmPaths {
    internal static string DirectoryPath(string name) {
        if (name.IndexOf('\0') >= 0 || name.Any(char.IsControl)) throw ChmBinary.Error("PATH", "A CHM path contains control characters.");
        if (name.StartsWith("::", StringComparison.Ordinal)) return name;
        string path = name.Replace('\\', '/');
        if (!path.StartsWith("/", StringComparison.Ordinal)) path = "/" + path;
        foreach (string component in path.Split('/')) {
            if (component == "." || component == ".." || component.IndexOf(':') >= 0)
                throw ChmBinary.Error("PATH", "A CHM entry contains an unsafe or ambiguous path: " + name);
        }
        if (path.Contains("//")) throw ChmBinary.Error("PATH", "A CHM entry contains empty path components.");
        return path;
    }

    internal static string? Resolve(string reference, string sourcePath, int maxLength) {
        if (reference.Length > maxLength || reference.Any(char.IsControl)) return null;
        string value = reference.Trim().Replace('\\', '/');
        if (value.IndexOf("::", StringComparison.Ordinal) >= 0 || value.StartsWith("//", StringComparison.Ordinal)) return null;
        int suffix = value.IndexOfAny(new[] { '#', '?' });
        string path = suffix >= 0 ? value.Substring(0, suffix) : value;
        if (path.IndexOf(':') >= 0) return null;
        try { path = Uri.UnescapeDataString(path); } catch (UriFormatException) { return null; }
        if (path.Any(char.IsControl) || path.IndexOfAny(new[] { ':', '\\' }) >= 0) return null;
        if (path.Length == 0) path = sourcePath;
        else if (!path.StartsWith("/", StringComparison.Ordinal)) path = sourcePath.Substring(0, sourcePath.LastIndexOf('/') + 1) + path;
        var components = new List<string>();
        foreach (string part in path.Split('/')) {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..") { if (components.Count == 0) return null; components.RemoveAt(components.Count - 1); }
            else components.Add(part);
        }
        return "/" + string.Join("/", components);
    }
}
