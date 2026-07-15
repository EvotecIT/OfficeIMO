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
        var pending = new Stack<string>();
        long totalBytes = 0;
        pending.Push(Path.GetFullPath(rootPath));

        while (pending.Count > 0 && results.Count < maxFiles) {
            cancellationToken.ThrowIfCancellationRequested();
            string directory = pending.Pop();

            string[] files = Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly);
            Array.Sort(files, StringComparer.Ordinal);
            foreach (string file in files) {
                cancellationToken.ThrowIfCancellationRequested();
                if (results.Count >= maxFiles) break;
                if (!extensions.Contains(Path.GetExtension(file)) &&
                    !string.Equals(Path.GetFileName(file), "winmail.dat", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                long length = new FileInfo(file).Length;
                if (maxTotalBytes.HasValue && totalBytes + length > maxTotalBytes.Value) {
                    return results;
                }

                totalBytes += length;
                results.Add(file);
            }

            if (!recurse) continue;
            string[] directories = Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly);
            Array.Sort(directories, StringComparer.Ordinal);
            for (int index = directories.Length - 1; index >= 0; index--) {
                FileAttributes attributes = File.GetAttributes(directories[index]);
                if ((attributes & FileAttributes.ReparsePoint) == 0) {
                    pending.Push(directories[index]);
                }
            }
        }

        return results;
    }
}