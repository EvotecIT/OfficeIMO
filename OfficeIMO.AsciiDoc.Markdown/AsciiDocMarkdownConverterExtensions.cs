namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Conversion extensions for native AsciiDoc documents.</summary>
public static class AsciiDocMarkdownConverterExtensions {
    /// <summary>Converts a native AsciiDoc document to the OfficeIMO Markdown semantic model.</summary>
    public static AsciiDocMarkdownConversionResult ToMarkdownDocument(
        this AsciiDocDocument document,
        AsciiDocToMarkdownOptions? options = null) =>
        AsciiDocToMarkdownConverter.Convert(document, options);

    /// <summary>Converts an OfficeIMO Markdown document to canonical lossless AsciiDoc.</summary>
    public static MarkdownAsciiDocConversionResult ToAsciiDocDocument(
        this MarkdownDoc document,
        MarkdownToAsciiDocOptions? options = null) =>
        MarkdownToAsciiDocConverter.Convert(document, options);
}
