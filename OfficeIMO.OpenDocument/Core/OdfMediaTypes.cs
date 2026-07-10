namespace OfficeIMO.OpenDocument;

/// <summary>Standard package media types for the supported OpenDocument family.</summary>
public static class OdfMediaTypes {
    /// <summary>ODT package media type.</summary>
    public const string Text = "application/vnd.oasis.opendocument.text";
    /// <summary>ODS package media type.</summary>
    public const string Spreadsheet = "application/vnd.oasis.opendocument.spreadsheet";
    /// <summary>ODP package media type.</summary>
    public const string Presentation = "application/vnd.oasis.opendocument.presentation";

    internal static string ForKind(OdfDocumentKind kind) {
        switch (kind) {
            case OdfDocumentKind.Text: return Text;
            case OdfDocumentKind.Spreadsheet: return Spreadsheet;
            case OdfDocumentKind.Presentation: return Presentation;
            default: throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }

    internal static bool TryGetKind(string mediaType, out OdfDocumentKind kind) {
        if (string.Equals(mediaType, Text, StringComparison.Ordinal)) { kind = OdfDocumentKind.Text; return true; }
        if (string.Equals(mediaType, Spreadsheet, StringComparison.Ordinal)) { kind = OdfDocumentKind.Spreadsheet; return true; }
        if (string.Equals(mediaType, Presentation, StringComparison.Ordinal)) { kind = OdfDocumentKind.Presentation; return true; }
        kind = OdfDocumentKind.Text;
        return false;
    }
}
