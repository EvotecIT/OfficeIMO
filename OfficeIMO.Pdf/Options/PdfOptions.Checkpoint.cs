namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>Font mappings whose program bytes are part of a conversion checkpoint's rendering identity.</summary>
    internal IReadOnlyList<PdfEmbeddedFont> CheckpointEmbeddedFonts =>
        _embeddedFonts?.OrderBy(pair => pair.Key).Select(pair => pair.Value).ToArray() ?? [];

    /// <summary>Distinguishes implicit defaults from explicitly assigned values used by adapter inheritance.</summary>
    internal string CheckpointExplicitSettings => string.Join(",", _hasExplicitDefaultFont, _hasExplicitHeaderFont,
        _hasExplicitFooterFont, _hasExplicitPageNumberStart, _hasExplicitDefaultTableStyle);
}
