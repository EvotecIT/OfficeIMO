using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.OneNote.Markdown;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.OneNote.Html;

/// <summary>Converts typed offline OneNote models to HTML through the shared Markdown model.</summary>
public static partial class OneNoteHtmlConverterExtensions {
    private static MarkdownHtmlExportOptions WholeDocument(MarkdownHtmlExportOptions? options) {
        var copy = options?.CloneForRender() ?? new MarkdownHtmlExportOptions();
        copy.Kind = MarkdownHtmlKind.Document;
        return copy;
    }
    internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this OneNoteSection section, OneNoteHtmlExportOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        PreparedHtml projection = Prepare(section, options?.ProjectionOptions);
        return new HtmlTextConversionResult(Render(projection, options), CreateHtmlDiagnostics(projection, section));
    }

    internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this OneNoteNotebook notebook, OneNoteHtmlExportOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        PreparedHtml projection = Prepare(notebook, options?.ProjectionOptions);
        return new HtmlTextConversionResult(Render(projection, options), CreateHtmlDiagnostics(projection, notebook));
    }

    internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this OneNoteSection section,
        OneNoteMarkdownExportOptions? projectionOptions = null, MarkdownHtmlExportOptions? htmlOptions = null, bool fragment = false) =>
        section.ConvertToHtmlWithEvidence(new OneNoteHtmlExportOptions { ProjectionOptions = projectionOptions, HtmlOptions = htmlOptions, Fragment = fragment });

    internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this OneNoteNotebook notebook,
        OneNoteMarkdownExportOptions? projectionOptions = null, MarkdownHtmlExportOptions? htmlOptions = null, bool fragment = false) =>
        notebook.ConvertToHtmlWithEvidence(new OneNoteHtmlExportOptions { ProjectionOptions = projectionOptions, HtmlOptions = htmlOptions, Fragment = fragment });

    private static string Render(PreparedHtml projection, OneNoteHtmlExportOptions? options) =>
        projection.Document.ToHtml(options?.Fragment == true ? options.HtmlOptions : WholeDocument(options?.HtmlOptions));

    private static byte[] Utf8(string value) => new UTF8Encoding(false).GetBytes(value);

    private static PreparedHtml Prepare(OneNoteSection section, OneNoteMarkdownExportOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteMarkdownExportOptions operation = CreateCachedOptions(options);
        OneNoteMarkdownConversionResult projection = section.ConvertToMarkdownWithEvidence(operation);
        return new PreparedHtml(OneNoteSemanticHtmlRenderer.CreateDocument(section, operation), projection, operation);
    }

    private static PreparedHtml Prepare(OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteMarkdownExportOptions operation = CreateCachedOptions(options);
        OneNoteMarkdownConversionResult projection = notebook.ConvertToMarkdownWithEvidence(operation);
        return new PreparedHtml(OneNoteSemanticHtmlRenderer.CreateDocument(notebook, operation), projection, operation);
    }

    private static OneNoteMarkdownExportOptions CreateCachedOptions(OneNoteMarkdownExportOptions? options) {
        OneNoteMarkdownExportOptions operation = (options ?? new OneNoteMarkdownExportOptions()).Clone();
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

    private static IEnumerable<HtmlDiagnostic> CreateHtmlDiagnostics(PreparedHtml projection, OneNoteSection section) =>
        projection.Projection.Diagnostics
            .Where(diagnostic => diagnostic.Code != "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED")
            .Select(ToHtmlDiagnostic)
            .Concat(OneNoteHtmlDiagnosticCollector.Collect(section, projection.Options));

    private static IEnumerable<HtmlDiagnostic> CreateHtmlDiagnostics(PreparedHtml projection, OneNoteNotebook notebook) =>
        projection.Projection.Diagnostics
            .Where(diagnostic => diagnostic.Code != "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED")
            .Select(ToHtmlDiagnostic)
            .Concat(OneNoteHtmlDiagnosticCollector.Collect(notebook, projection.Options));

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
        internal PreparedHtml(MarkdownDocument document, OneNoteMarkdownConversionResult projection, OneNoteMarkdownExportOptions options) {
            Document = document;
            Projection = projection;
            Options = options;
        }

        internal MarkdownDocument Document { get; }
        internal OneNoteMarkdownConversionResult Projection { get; }
        internal OneNoteMarkdownExportOptions Options { get; }
    }
}
