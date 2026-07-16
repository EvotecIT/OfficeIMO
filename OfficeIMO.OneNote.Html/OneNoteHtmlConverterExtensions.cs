using OfficeIMO.Markdown;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Html;

/// <summary>Converts typed offline OneNote models to HTML through the shared Markdown model.</summary>
public static class OneNoteHtmlConverterExtensions {
    /// <summary>Converts a section to a standalone HTML5 document.</summary>
    public static string ToHtmlDocument(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).ToHtmlDocument(htmlOptions);

    /// <summary>Converts a notebook to a standalone HTML5 document.</summary>
    public static string ToHtmlDocument(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).ToHtmlDocument(htmlOptions);

    /// <summary>Converts a section to an embeddable HTML fragment.</summary>
    public static string ToHtmlFragment(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).ToHtmlFragment(htmlOptions);

    /// <summary>Converts a notebook to an embeddable HTML fragment.</summary>
    public static string ToHtmlFragment(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).ToHtmlFragment(htmlOptions);

    /// <summary>Encodes a section as a standalone UTF-8 HTML document without a byte-order mark.</summary>
    public static byte[] ToHtmlBytes(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Utf8(section.ToHtmlDocument(projectionOptions, htmlOptions));

    /// <summary>Encodes a notebook as a standalone UTF-8 HTML document without a byte-order mark.</summary>
    public static byte[] ToHtmlBytes(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Utf8(notebook.ToHtmlDocument(projectionOptions, htmlOptions));

    /// <summary>Saves a section as a standalone HTML document.</summary>
    public static void SaveAsHtml(
        this OneNoteSection section,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsHtml(path, htmlOptions);

    /// <summary>Saves a notebook as a standalone HTML document.</summary>
    public static void SaveAsHtml(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsHtml(path, htmlOptions);

    /// <summary>Writes a section as a standalone HTML document to a caller-owned stream.</summary>
    public static void SaveAsHtml(
        this OneNoteSection section,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Write(stream, section.ToHtmlBytes(projectionOptions, htmlOptions));

    /// <summary>Writes a notebook as a standalone HTML document to a caller-owned stream.</summary>
    public static void SaveAsHtml(
        this OneNoteNotebook notebook,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Write(stream, notebook.ToHtmlBytes(projectionOptions, htmlOptions));

    /// <summary>Asynchronously saves a section as a standalone HTML document.</summary>
    public static Task SaveAsHtmlAsync(
        this OneNoteSection section,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null,
        CancellationToken cancellationToken = default) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsHtmlAsync(path, htmlOptions, cancellationToken);

    /// <summary>Asynchronously saves a notebook as a standalone HTML document.</summary>
    public static Task SaveAsHtmlAsync(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null,
        CancellationToken cancellationToken = default) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsHtmlAsync(path, htmlOptions, cancellationToken);

    /// <summary>Asynchronously writes a section as HTML to a caller-owned stream.</summary>
    public static Task SaveAsHtmlAsync(
        this OneNoteSection section,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null,
        CancellationToken cancellationToken = default) =>
        WriteAsync(stream, section.ToHtmlBytes(projectionOptions, htmlOptions), cancellationToken);

    /// <summary>Asynchronously writes a notebook as HTML to a caller-owned stream.</summary>
    public static Task SaveAsHtmlAsync(
        this OneNoteNotebook notebook,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null,
        CancellationToken cancellationToken = default) =>
        WriteAsync(stream, notebook.ToHtmlBytes(projectionOptions, htmlOptions), cancellationToken);

    private static byte[] Utf8(string value) => new UTF8Encoding(false).GetBytes(value);

    private static void Write(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        stream.Write(bytes, 0, bytes.Length);
    }

    private static Task WriteAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken);
    }
}
