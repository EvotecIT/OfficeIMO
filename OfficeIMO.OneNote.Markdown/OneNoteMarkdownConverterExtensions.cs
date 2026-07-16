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
        MarkdownReader.Parse(section.ToMarkdown(options));

    /// <summary>Converts a notebook to a first-party Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) =>
        MarkdownReader.Parse(notebook.ToMarkdown(options));

    /// <summary>Encodes section Markdown as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToMarkdownBytes(this OneNoteSection section, OneNoteMarkdownOptions? options = null) =>
        new UTF8Encoding(false).GetBytes(section.ToMarkdown(options));

    /// <summary>Encodes notebook Markdown as UTF-8 without a byte-order mark.</summary>
    public static byte[] ToMarkdownBytes(this OneNoteNotebook notebook, OneNoteMarkdownOptions? options = null) =>
        new UTF8Encoding(false).GetBytes(notebook.ToMarkdown(options));
}
