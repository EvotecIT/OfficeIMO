using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static string? NormalizeMetadata(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string normalized = value!.Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static string? FindMatchingFirstHeadingAnchor(MarkdownDoc document, string title) {
        for (int i = 0; i < document.TopLevelBlocks.Count; i++) {
            if (document.TopLevelBlocks[i] is HeadingBlock heading && heading.Level == 1 && IsSameNormalizedText(heading.Text, title)) {
                return document.GetHeadingAnchor(heading);
            }
        }

        return null;
    }

    private static bool IsSameNormalizedText(string? left, string? right) {
        string? normalizedLeft = NormalizeMetadata(left);
        string? normalizedRight = NormalizeMetadata(right);
        return normalizedLeft != null
            && normalizedRight != null
            && string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatTitleFromKind(string? kind) {
        if (string.IsNullOrWhiteSpace(kind)) {
            return "Note";
        }

        string trimmed = kind!.Trim();
        return trimmed.Length == 1
            ? trimmed.ToUpperInvariant()
            : char.ToUpperInvariant(trimmed[0]) + trimmed.Substring(1).ToLowerInvariant();
    }


    private static bool TryGetBookmarkTarget(string? value, out string? bookmarkName) {
        bookmarkName = null;
        if (string.IsNullOrWhiteSpace(value) || !value!.StartsWith("#", StringComparison.Ordinal) || value.Length == 1) {
            return false;
        }

        bookmarkName = value.Substring(1);
        return !string.IsNullOrWhiteSpace(bookmarkName);
    }

    private static string? NormalizeAbsoluteLink(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeMailto)
            ? uri.ToString()
            : null;
    }
}
