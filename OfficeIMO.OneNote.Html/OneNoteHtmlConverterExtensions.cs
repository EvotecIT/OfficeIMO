using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.OneNote.Markdown;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.OneNote.Html;

/// <summary>Converts typed offline OneNote models to HTML through the shared Markdown model.</summary>
public static class OneNoteHtmlConverterExtensions {
    /// <summary>Converts a section to a standalone HTML5 document.</summary>
    public static string ToHtmlDocument(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Prepare(section, projectionOptions).Document.ToHtmlDocument(htmlOptions);

    /// <summary>Converts a section to HTML with shared structured projection diagnostics.</summary>
    public static HtmlTextConversionResult ToHtmlDocumentResult(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        PreparedHtml projection = Prepare(section, projectionOptions);
        return new HtmlTextConversionResult(
            projection.Document.ToHtmlDocument(htmlOptions),
            projection.Projection.Diagnostics
                .Where(diagnostic => diagnostic.Code != "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED")
                .Select(ToHtmlDiagnostic));
    }

    /// <summary>Converts a notebook to a standalone HTML5 document.</summary>
    public static string ToHtmlDocument(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Prepare(notebook, projectionOptions).Document.ToHtmlDocument(htmlOptions);

    /// <summary>Converts a notebook to HTML with shared structured projection diagnostics.</summary>
    public static HtmlTextConversionResult ToHtmlDocumentResult(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        PreparedHtml projection = Prepare(notebook, projectionOptions);
        return new HtmlTextConversionResult(
            projection.Document.ToHtmlDocument(htmlOptions),
            projection.Projection.Diagnostics
                .Where(diagnostic => diagnostic.Code != "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED")
                .Select(ToHtmlDiagnostic));
    }

    /// <summary>Converts a section to an embeddable HTML fragment.</summary>
    public static string ToHtmlFragment(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Prepare(section, projectionOptions).Document.ToHtmlFragment(htmlOptions);

    /// <summary>Converts a notebook to an embeddable HTML fragment.</summary>
    public static string ToHtmlFragment(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Prepare(notebook, projectionOptions).Document.ToHtmlFragment(htmlOptions);

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
        Prepare(section, projectionOptions).Document.SaveAsHtml(path, htmlOptions);

    /// <summary>Saves a notebook as a standalone HTML document.</summary>
    public static void SaveAsHtml(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null) =>
        Prepare(notebook, projectionOptions).Document.SaveAsHtml(path, htmlOptions);

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
        Prepare(section, projectionOptions).Document.SaveAsHtmlAsync(path, htmlOptions, cancellationToken);

    /// <summary>Asynchronously saves a notebook as a standalone HTML document.</summary>
    public static Task SaveAsHtmlAsync(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        HtmlOptions? htmlOptions = null,
        CancellationToken cancellationToken = default) =>
        Prepare(notebook, projectionOptions).Document.SaveAsHtmlAsync(path, htmlOptions, cancellationToken);

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

    private static PreparedHtml Prepare(OneNoteSection section, OneNoteMarkdownOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteMarkdownOptions operation = CreateCachedOptions(options);
        OneNoteMarkdownConversionResult projection = section.ToMarkdownDocumentResult(operation);
        return new PreparedHtml(OneNoteSemanticHtmlRenderer.CreateDocument(section, operation), projection);
    }

    private static PreparedHtml Prepare(OneNoteNotebook notebook, OneNoteMarkdownOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteMarkdownOptions operation = CreateCachedOptions(options);
        OneNoteMarkdownConversionResult projection = notebook.ToMarkdownDocumentResult(operation);
        return new PreparedHtml(OneNoteSemanticHtmlRenderer.CreateDocument(notebook, operation), projection);
    }

    private static OneNoteMarkdownOptions CreateCachedOptions(OneNoteMarkdownOptions? options) {
        OneNoteMarkdownOptions operation = (options ?? new OneNoteMarkdownOptions()).Clone();
        Func<OneNoteBinaryElement, string?>? resolver = operation.AssetUriResolver;
        if (resolver == null) return operation;
        var cache = new Dictionary<OneNoteBinaryElement, string?>();
        operation.AssetUriResolver = element => {
            if (cache.TryGetValue(element, out string? value)) return value;
            value = resolver(element);
            cache[element] = value;
            return value;
        };
        return operation;
    }

    private static HtmlDiagnostic ToHtmlDiagnostic(OneNoteMarkdownDiagnostic diagnostic) {
        HtmlDiagnosticSeverity severity = diagnostic.Severity == OneNoteDiagnosticSeverity.Error
            ? HtmlDiagnosticSeverity.Error
            : diagnostic.Severity == OneNoteDiagnosticSeverity.Warning
                ? HtmlDiagnosticSeverity.Warning
                : HtmlDiagnosticSeverity.Info;
        OfficeConversionLossKind lossKind = severity == HtmlDiagnosticSeverity.Error
            ? OfficeConversionLossKind.Failure
            : severity == HtmlDiagnosticSeverity.Warning
                ? OfficeConversionLossKind.Approximation
                : OfficeConversionLossKind.None;
        return new HtmlDiagnostic(
            "OfficeIMO.OneNote.Html",
            diagnostic.Code,
            diagnostic.Message,
            severity,
            diagnostic.Source,
            lossKind: lossKind);
    }

    private static void Write(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        stream.Write(bytes, 0, bytes.Length);
    }

    private static Task WriteAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken);
    }

    private sealed class PreparedHtml {
        internal PreparedHtml(MarkdownDoc document, OneNoteMarkdownConversionResult projection) {
            Document = document;
            Projection = projection;
        }

        internal MarkdownDoc Document { get; }
        internal OneNoteMarkdownConversionResult Projection { get; }
    }
}
