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
        var results = new List<string>(Math.Min(maxFiles, 256));
        long totalBytes = 0;
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

            long length = new FileInfo(file).Length;
            if (maxTotalBytes.HasValue && length > maxTotalBytes.Value - totalBytes) {
                break;
            }

            totalBytes += length;
            results.Add(file);
            if (results.Count >= maxFiles) break;
        }

        results.Sort(StringComparer.Ordinal);
        return results;
    }
}
