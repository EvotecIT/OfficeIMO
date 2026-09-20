using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

/// <summary>Shared font programs and script runs used by raster, SVG, PDF, and HarfBuzz evidence.</summary>
public static class TypographyEvidenceCorpus {
    /// <summary>Gets the representative multilingual script matrix.</summary>
    public static IReadOnlyList<TypographyEvidenceCase> Cases { get; } = new[] {
        new TypographyEvidenceCase("GreekCyrillicFallback", "Carlito", "Carlito-Regular.ttf", "Ελληνικά Кириллица", OfficeTextDirection.LeftToRight, "el", false, true),
        new TypographyEvidenceCase("Arabic", "Noto Sans Arabic", "NotoSansArabic-Regular.ttf", "سلام بالعالم", OfficeTextDirection.RightToLeft, "ar", true, true),
        new TypographyEvidenceCase("Hebrew", "OfficeIMO Hebrew Evidence", "", "שלום עולם", OfficeTextDirection.RightToLeft, "he", true, false),
        new TypographyEvidenceCase("Indic", "Noto Sans Devanagari", "NotoSansDevanagari-Regular.ttf", "नमस्ते दुनिया", OfficeTextDirection.LeftToRight, "hi", false, true),
        new TypographyEvidenceCase("CjkVertical", "Noto Sans SC", "NotoSansSC-BaselineSubset.ttf", "永字国", OfficeTextDirection.TopToBottom, "zh", false, false),
        new TypographyEvidenceCase("CombiningMarks", "OfficeIMO Combining Evidence", "", "Cafe\u0301", OfficeTextDirection.LeftToRight, "fr", false, false),
        new TypographyEvidenceCase("Emoji", "Noto Emoji", "NotoEmoji-VariableFont_wght.ttf", "😀🚀🌍", OfficeTextDirection.LeftToRight, "und", false, true)
    };
}

/// <summary>One deterministic typography evidence case.</summary>
public sealed class TypographyEvidenceCase {
    /// <summary>Creates a typography evidence case.</summary>
    public TypographyEvidenceCase(
        string name,
        string family,
        string fontFileName,
        string text,
        OfficeTextDirection direction,
        string language,
        bool managedShapingExpected,
        bool harfBuzzShapingExpected) {
        Name = name;
        Family = family;
        FontFileName = fontFileName;
        Text = text;
        Direction = direction;
        Language = language;
        ManagedShapingExpected = managedShapingExpected;
        HarfBuzzShapingExpected = harfBuzzShapingExpected;
    }

    /// <summary>Stable case name.</summary>
    public string Name { get; }
    /// <summary>Font family used by every renderer.</summary>
    public string Family { get; }
    /// <summary>Exact font-program file name.</summary>
    public string FontFileName { get; }
    /// <summary>Logical Unicode text.</summary>
    public string Text { get; }
    /// <summary>Shaping direction, including vertical CJK evidence.</summary>
    public OfficeTextDirection Direction { get; }
    /// <summary>BCP 47 language hint.</summary>
    public string Language { get; }
    /// <summary>Whether the bounded managed shaper is expected to accept the run.</summary>
    public bool ManagedShapingExpected { get; }
    /// <summary>Whether HarfBuzz can shape this repository-owned font fixture.</summary>
    public bool HarfBuzzShapingExpected { get; }
}
