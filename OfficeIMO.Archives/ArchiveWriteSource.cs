namespace OfficeIMO.Archives;

/// <summary>Enumerates a source folder into archive entries in ordinal path order.</summary>
internal static class ArchiveWriteSource {
    internal static List<ArchiveWriteEntry> FromDirectory(string sourceDirectory, ArchiveWriterOptions options, CancellationToken token) {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceDirectory);
        string root = Path.GetFullPath(sourceDirectory);
        if (!Directory.Exists(root)) throw new DirectoryNotFoundException("The source directory does not exist: " + sourceDirectory);
        var entries = new List<(string Path, ArchiveWriteEntry Entry)>();
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0) {
            token.ThrowIfCancellationRequested();
            string directory = pending.Pop();
            foreach (string child in Directory.EnumerateFileSystemEntries(directory)) {
                token.ThrowIfCancellationRequested();
                var info = new FileInfo(child);
                if ((info.Attributes & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Archive sources must not contain symbolic links or reparse points: " + child);
                string relative = child.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    .Replace(Path.DirectorySeparatorChar, '/');
                DateTimeOffset? time = options.PreserveFileTimes ? new DateTimeOffset(info.LastWriteTimeUtc, TimeSpan.Zero) : null;
                if ((info.Attributes & FileAttributes.Directory) != 0) {
                    entries.Add((relative + "/", ArchiveWriteEntry.Directory(relative, time)));
                    pending.Push(child);
                } else {
                    string file = child;
                    entries.Add((relative, ArchiveWriteEntry.File(relative,
                        () => new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read), time)));
                }
            }
        }
        return entries.OrderBy(item => item.Path, StringComparer.Ordinal).Select(item => item.Entry).ToList();
    }

    internal static void WriteToFile(string destinationPath, Action<Stream> write) {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        string full = Path.GetFullPath(destinationPath);
        string? parent = Path.GetDirectoryName(full);
        if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
        var output = new FileStream(full, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        bool completed = false;
        try {
            write(output);
            output.Dispose();
            completed = true;
        } finally {
            if (!completed) {
                output.Dispose();
                try { File.Delete(full); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
    }
}
