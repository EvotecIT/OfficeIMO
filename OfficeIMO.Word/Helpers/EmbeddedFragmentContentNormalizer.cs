using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

/// <summary>
/// Applies format-specific normalization before an alternative-format fragment is embedded.
/// </summary>
internal static class EmbeddedFragmentContentNormalizer {
    internal static string Normalize(
        MainDocumentPart mainDocumentPart,
        string content,
        WordAlternativeFormatImportPartType partType) {
        return partType switch {
            WordAlternativeFormatImportPartType.Html => HtmlListSemanticsNormalizer.Normalize(content),
            WordAlternativeFormatImportPartType.Rtf => RtfListSemanticsNormalizer.Normalize(content, mainDocumentPart),
            _ => content
        };
    }
}
