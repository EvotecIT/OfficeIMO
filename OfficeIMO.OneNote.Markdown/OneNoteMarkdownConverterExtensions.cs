using OfficeIMO.Markdown;

namespace OfficeIMO.OneNote.Markdown;

/// <summary>Converts typed offline OneNote models through the first-party Markdown model.</summary>
public static partial class OneNoteMarkdownConverterExtensions {
    internal static OneNoteMarkdownConversionResult ConvertToMarkdownWithEvidence(this OneNoteSection section, OneNoteMarkdownExportOptions? options) =>
        ConvertToMarkdownWithEvidence(section, options, out _);

    internal static OneNoteMarkdownConversionResult ConvertToMarkdownWithEvidence(this OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options) =>
        ConvertToMarkdownWithEvidence(notebook, options, out _);

    private static OneNoteMarkdownConversionResult ConvertToMarkdownWithEvidence(OneNoteSection section, OneNoteMarkdownExportOptions? options, out string text) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteMarkdownExportOptions operation = CreateTrackedOptions(options, out HashSet<OneNoteBinaryElement> resolvedAssets);
        text = OneNoteMarkdownProjection.ToMarkdown(section, operation);
        return new OneNoteMarkdownConversionResult(MarkdownDocument.Parse(text), OneNoteMarkdownDiagnosticCollector.Collect(section, operation, resolvedAssets));
    }

    private static OneNoteMarkdownConversionResult ConvertToMarkdownWithEvidence(OneNoteNotebook notebook, OneNoteMarkdownExportOptions? options, out string text) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteMarkdownExportOptions operation = CreateTrackedOptions(options, out HashSet<OneNoteBinaryElement> resolvedAssets);
        text = OneNoteMarkdownProjection.ToMarkdown(notebook, operation);
        return new OneNoteMarkdownConversionResult(MarkdownDocument.Parse(text), OneNoteMarkdownDiagnosticCollector.Collect(notebook, operation, resolvedAssets));
    }

    private static OneNoteMarkdownExportOptions CreateTrackedOptions(
        OneNoteMarkdownExportOptions? options,
        out HashSet<OneNoteBinaryElement> resolvedAssets) {
        OneNoteMarkdownExportOptions operation = (options ?? new OneNoteMarkdownExportOptions()).Clone();
        resolvedAssets = new HashSet<OneNoteBinaryElement>();
        Func<OneNoteBinaryElement, string?>? resolver = operation.AssetUriResolver;
        if (resolver == null) return operation;

        HashSet<OneNoteBinaryElement> outcomes = resolvedAssets;
        operation.AssetUriResolver = element => {
            string? uri = resolver(element);
            if (uri != null) outcomes.Add(element);
            return uri;
        };
        return operation;
    }
}
