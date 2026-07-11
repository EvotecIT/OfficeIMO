namespace OfficeIMO.Drawing;

/// <summary>One grapheme-safe text segment assigned to a resolved font-family fallback.</summary>
public sealed class OfficeFontFallbackRun {
    internal OfficeFontFallbackRun(string text, string familyName) {
        Text = text;
        FamilyName = familyName;
    }

    /// <summary>Text retained in logical source order.</summary>
    public string Text { get; }

    /// <summary>Resolved family name, or the original family list when no scoped face covers the segment.</summary>
    public string FamilyName { get; }
}
