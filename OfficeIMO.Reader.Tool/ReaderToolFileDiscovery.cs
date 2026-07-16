namespace OfficeIMO.Reader.Tool;

internal static class ReaderToolFileDiscovery {
    internal static IReadOnlyList<string> FindSupportedFiles(
        string rootPath,
        OfficeDocumentReader reader,
        bool recurse,
        int maxFiles,
        long? maxTotalBytes,
        CancellationToken cancellationToken) {
        var extensions = new HashSet<string>(
            reader.GetCapabilities().SelectMany(capability => capability.Extensions),
            StringComparer.OrdinalIgnoreCase);
        var selected = new SortedSet<string>(StringComparer.Ordinal);
        var enumerationOptions = new EnumerationOptions {
            RecurseSubdirectories = recurse,
            AttributesToSkip = FileAttributes.ReparsePoint,
            ReturnSpecialDirectories = false
        };

        foreach (string file in Directory.EnumerateFiles(
            Path.GetFullPath(rootPath),
            "*",
            enumerationOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!extensions.Contains(Path.GetExtension(file)) &&
                !string.Equals(Path.GetFileName(file), "winmail.dat", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            selected.Add(file);
            if (selected.Count > maxFiles) {
                selected.Remove(selected.Max!);
            }
        }

        var results = new List<string>(selected.Count);
        long totalBytes = 0;
        foreach (string file in selected) {
            long length = new FileInfo(file).Length;
            if (maxTotalBytes.HasValue && length > maxTotalBytes.Value - totalBytes) break;
            totalBytes += length;
            results.Add(file);
        }
        return results;
    }
}
