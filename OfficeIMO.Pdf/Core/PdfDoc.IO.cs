using System;
using System.IO;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>
    /// Renders the document into a PDF byte array in memory.
    /// </summary>
    public byte[] ToBytes() => PdfWriter.Write(this, _blocks, _options, _title, _author, _subject, _keywords);

    /// <summary>
    /// Saves the document to <paramref name="path"/>. Creates the directory if needed.
    /// </summary>
    /// <param name="path">Destination file path, e.g. "C:\\Docs\\Report.pdf".</param>
    /// <returns>This <see cref="PdfDoc"/> for chaining.</returns>
    public PdfDoc Save(string path) {
        Guard.NotNull(path, nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));

        string fullPath;
        try { fullPath = Path.GetFullPath(path); }
        catch (Exception ex) { throw new ArgumentException("Path is invalid.", nameof(path), ex); }

        // Reject if points to an existing directory or a root without filename
        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory)
            throw new ArgumentException("Path refers to a directory; a file path is required.", nameof(path));

        // Validate file name characters proactively
        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) throw new ArgumentException("Path must include a file name.", nameof(path));
        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            throw new ArgumentException("Path contains invalid file name characters.", nameof(path));

        var dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(dir)) dir = ".";
        Directory.CreateDirectory(dir);

        var bytes = ToBytes();
        File.WriteAllBytes(fullPath, bytes);
        return this;
    }

    /// <summary>
    /// Asynchronously saves the document to <paramref name="path"/>.
    /// </summary>
    public async System.Threading.Tasks.Task SaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(path, nameof(path));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));

        string fullPath;
        try { fullPath = Path.GetFullPath(path); }
        catch (Exception ex) { throw new ArgumentException("Path is invalid.", nameof(path), ex); }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory)
            throw new ArgumentException("Path refers to a directory; a file path is required.", nameof(path));

        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) throw new ArgumentException("Path must include a file name.", nameof(path));
        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            throw new ArgumentException("Path contains invalid file name characters.", nameof(path));

        var dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(dir)) dir = ".";
        Directory.CreateDirectory(dir);

        var bytes = ToBytes();
        using var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
#if NET8_0_OR_GREATER
        await fs.WriteAsync(bytes.AsMemory(0, bytes.Length), cancellationToken).ConfigureAwait(false);
#else
        await fs.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }
}
