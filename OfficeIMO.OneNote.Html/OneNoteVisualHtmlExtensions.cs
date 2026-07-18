namespace OfficeIMO.OneNote.Html;

/// <summary>Visual-preserving HTML entry points backed by the native OneNote Drawing canvas.</summary>
public static class OneNoteVisualHtmlExtensions {
    /// <summary>Converts a section to a standalone responsive HTML5 document containing one SVG canvas per page.</summary>
    public static string ToVisualHtmlDocument(this OneNoteSection section, OneNoteVisualHtmlOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        return OneNoteVisualHtmlRenderer.RenderDocument(section.Name, OneNotePageTraversal.Flatten(section), options);
    }

    /// <summary>Converts a notebook to a standalone responsive HTML5 document containing one SVG canvas per page.</summary>
    public static string ToVisualHtmlDocument(this OneNoteNotebook notebook, OneNoteVisualHtmlOptions? options = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        return OneNoteVisualHtmlRenderer.RenderDocument(notebook.Name, OneNotePageTraversal.Flatten(notebook), options);
    }

    /// <summary>Converts a section to an embeddable responsive SVG-page fragment.</summary>
    public static string ToVisualHtmlFragment(this OneNoteSection section, OneNoteVisualHtmlOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        return OneNoteVisualHtmlRenderer.RenderFragment(OneNotePageTraversal.Flatten(section), options);
    }

    /// <summary>Converts a notebook to an embeddable responsive SVG-page fragment.</summary>
    public static string ToVisualHtmlFragment(this OneNoteNotebook notebook, OneNoteVisualHtmlOptions? options = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        return OneNoteVisualHtmlRenderer.RenderFragment(OneNotePageTraversal.Flatten(notebook), options);
    }

    /// <summary>Encodes a section's visual HTML document as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToVisualHtmlBytes(this OneNoteSection section, OneNoteVisualHtmlOptions? options = null) => Utf8(section.ToVisualHtmlDocument(options));

    /// <summary>Encodes a notebook's visual HTML document as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToVisualHtmlBytes(this OneNoteNotebook notebook, OneNoteVisualHtmlOptions? options = null) => Utf8(notebook.ToVisualHtmlDocument(options));

    /// <summary>Saves a section's visual HTML document.</summary>
    public static void SaveAsVisualHtml(this OneNoteSection section, string path, OneNoteVisualHtmlOptions? options = null) => WritePath(path, section.ToVisualHtmlBytes(options));

    /// <summary>Saves a notebook's visual HTML document.</summary>
    public static void SaveAsVisualHtml(this OneNoteNotebook notebook, string path, OneNoteVisualHtmlOptions? options = null) => WritePath(path, notebook.ToVisualHtmlBytes(options));

    /// <summary>Writes a section's visual HTML document to a caller-owned stream.</summary>
    public static void SaveAsVisualHtml(this OneNoteSection section, Stream stream, OneNoteVisualHtmlOptions? options = null) => Write(stream, section.ToVisualHtmlBytes(options));

    /// <summary>Writes a notebook's visual HTML document to a caller-owned stream.</summary>
    public static void SaveAsVisualHtml(this OneNoteNotebook notebook, Stream stream, OneNoteVisualHtmlOptions? options = null) => Write(stream, notebook.ToVisualHtmlBytes(options));

    /// <summary>Asynchronously saves a section's visual HTML document.</summary>
    public static Task SaveAsVisualHtmlAsync(this OneNoteSection section, string path, OneNoteVisualHtmlOptions? options = null, CancellationToken cancellationToken = default) =>
        WritePathAsync(path, section.ToVisualHtmlBytes(options), cancellationToken);

    /// <summary>Asynchronously saves a notebook's visual HTML document.</summary>
    public static Task SaveAsVisualHtmlAsync(this OneNoteNotebook notebook, string path, OneNoteVisualHtmlOptions? options = null, CancellationToken cancellationToken = default) =>
        WritePathAsync(path, notebook.ToVisualHtmlBytes(options), cancellationToken);

    /// <summary>Asynchronously writes a section's visual HTML document to a caller-owned stream.</summary>
    public static Task SaveAsVisualHtmlAsync(this OneNoteSection section, Stream stream, OneNoteVisualHtmlOptions? options = null, CancellationToken cancellationToken = default) =>
        WriteAsync(stream, section.ToVisualHtmlBytes(options), cancellationToken);

    /// <summary>Asynchronously writes a notebook's visual HTML document to a caller-owned stream.</summary>
    public static Task SaveAsVisualHtmlAsync(this OneNoteNotebook notebook, Stream stream, OneNoteVisualHtmlOptions? options = null, CancellationToken cancellationToken = default) =>
        WriteAsync(stream, notebook.ToVisualHtmlBytes(options), cancellationToken);

    private static byte[] Utf8(string value) => new UTF8Encoding(false).GetBytes(value);

    private static void WritePath(string path, byte[] bytes) {
        string fullPath = PreparePath(path);
        File.WriteAllBytes(fullPath, bytes);
    }

    private static async Task WritePathAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        string fullPath = PreparePath(path);
        using (var stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true)) {
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
        }
    }

    private static string PreparePath(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be empty.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        return fullPath;
    }

    private static void Write(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        stream.Write(bytes, 0, bytes.Length);
    }

    private static Task WriteAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken);
    }
}
