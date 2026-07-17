using OfficeIMO.Markdown;

namespace OfficeIMO.OneNote.Markdown;

/// <summary>Converts typed offline OneNote models through the first-party Markdown model.</summary>
public static class OneNoteMarkdownConverterExtensions {
    /// <summary>Converts a section to Markdown text.</summary>
    public static string ToMarkdown(this OneNoteSection section, OneNoteMarkdownOptions? options = null) =>
        OneNoteMarkdownProjection.ToMarkdown(section, options);

    /// <summary>Converts a notebook to Markdown text.</summary>
    public static string ToMarkdown(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) =>
        OneNoteMarkdownProjection.ToMarkdown(notebook, options);

    /// <summary>Converts a section to a first-party Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(this OneNoteSection section, OneNoteMarkdownOptions? options = null) =>
        section.ToMarkdownDocumentResult(options).Value;

    /// <summary>Converts a notebook to a first-party Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) =>
        notebook.ToMarkdownDocumentResult(options).Value;

    /// <summary>Converts a section to Markdown with explicit source and semantic-projection diagnostics.</summary>
    public static OneNoteMarkdownConversionResult ToMarkdownDocumentResult(this OneNoteSection section, OneNoteMarkdownOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteMarkdownOptions operation = options ?? new OneNoteMarkdownOptions();
        MarkdownDoc value = MarkdownReader.Parse(section.ToMarkdown(operation));
        return new OneNoteMarkdownConversionResult(value, OneNoteMarkdownDiagnosticCollector.Collect(section, operation));
    }

    /// <summary>Converts a notebook to Markdown with explicit source and semantic-projection diagnostics.</summary>
    public static OneNoteMarkdownConversionResult ToMarkdownDocumentResult(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteMarkdownOptions operation = options ?? new OneNoteMarkdownOptions();
        MarkdownDoc value = MarkdownReader.Parse(notebook.ToMarkdown(operation));
        return new OneNoteMarkdownConversionResult(value, OneNoteMarkdownDiagnosticCollector.Collect(notebook, operation));
    }

    /// <summary>Encodes section Markdown as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToMarkdownBytes(this OneNoteSection section, OneNoteMarkdownOptions? options = null) =>
        new UTF8Encoding(false).GetBytes(section.ToMarkdown(options));

    /// <summary>Encodes notebook Markdown as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToMarkdownBytes(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) =>
        new UTF8Encoding(false).GetBytes(notebook.ToMarkdown(options));
}
