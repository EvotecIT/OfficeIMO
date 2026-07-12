namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Conversion extensions for native AsciiDoc documents.</summary>
public static class AsciiDocMarkdownConverterExtensions {
    /// <summary>Converts a native AsciiDoc document to the OfficeIMO Markdown semantic model.</summary>
    public static AsciiDocToMarkdownResult ToMarkdownDocumentResult(
        this AsciiDocDocument document,
        AsciiDocToMarkdownOptions? options = null) =>
        AsciiDocToMarkdownConverter.Convert(document, options);

    /// <summary>Converts an AsciiDoc document to a typed Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(
        this AsciiDocDocument document,
        AsciiDocToMarkdownOptions? options = null) =>
        document.ToMarkdownDocumentResult(options).Value;

    /// <summary>Converts an OfficeIMO Markdown document to canonical lossless AsciiDoc.</summary>
    public static MarkdownToAsciiDocResult ToAsciiDocDocumentResult(
        this MarkdownDoc document,
        MarkdownToAsciiDocOptions? options = null) =>
        MarkdownToAsciiDocConverter.Convert(document, options);

    /// <summary>Converts a Markdown document to a parsed canonical AsciiDoc document.</summary>
    public static AsciiDocDocument ToAsciiDocDocument(
        this MarkdownDoc document,
        MarkdownToAsciiDocOptions? options = null) =>
        document.ToAsciiDocDocumentResult(options).Value;
}
