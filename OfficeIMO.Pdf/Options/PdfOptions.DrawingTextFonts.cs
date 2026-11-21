using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    /// <summary>
    /// Prepares fonts for positioned drawing scenes written as real PDF text: embeds the installed
    /// families the scenes name and, when their text leaves WinAnsi, registers Unicode fallbacks
    /// chosen by coverage of that text. Families owned by a rendering profile or already registered stay as configured.
    /// </summary>
    internal PdfCreateOptions UseDrawingTextFonts(IEnumerable<OfficeDrawing> drawings, CancellationToken cancellationToken = default) {
        Guard.NotNull(drawings, nameof(drawings));
        var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var unicodeText = new System.Text.StringBuilder();
        var seenElements = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeDrawing drawing in drawings) {
            CollectDrawingText(drawing, families, unicodeText, seenElements, cancellationToken);
        }

        foreach (string family in families) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasRenderingProfileFamilyPlanner(family) || HasNamedFontFamily(family)) continue;
            TryRegisterNamedOfficeFontFamily(family, out _);
        }

        if (unicodeText.Length != 0) {
            UseTextFallbacks(
                PdfTextFallbackFeatures.DocumentFont | PdfTextFallbackFeatures.SymbolAndEmojiFonts | PdfTextFallbackFeatures.MultilingualFonts,
                Array.Empty<PdfStandardFont>(),
                allowSystemFontEmbedding: true,
                preserveConfiguredFontSlots: HasExplicitDefaultFont,
                requiredText: unicodeText.ToString());
        }
        return this;
    }

    private Dictionary<string, bool>? _drawingTextCoverage;

    /// <summary>Reports whether any drawing text has a character that no configured font, fallback or standard encoding can write.</summary>
    internal bool HasUncoveredDrawingText(IEnumerable<OfficeDrawing> drawings, CancellationToken cancellationToken = default) {
        Guard.NotNull(drawings, nameof(drawings));
        foreach (OfficeDrawing drawing in drawings) {
            if (HasUncoveredDrawingText(drawing, cancellationToken)) return true;
        }
        return false;
    }

    /// <summary>
    /// Returns drawing text that the configured fonts can write. Each uncovered text element is replaced by its
    /// Unicode compatibility form when that is covered (for example mathematical italic a becomes a), otherwise by
    /// a visible missing-glyph mark. <paramref name="onSubstitution"/> receives each original and replacement.
    /// </summary>
    internal string SubstituteUncoveredDrawingText(string text, string? familyName, bool bold, bool italic, Action<string, string>? onSubstitution = null) {
        if (string.IsNullOrEmpty(text) || !PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(text) ||
            IsDrawingTextCovered(text, familyName, bold, italic)) return text;
        var builder = new System.Text.StringBuilder(text.Length);
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            if (!PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(element) || IsDrawingTextCovered(element, familyName, bold, italic)) {
                builder.Append(element);
                continue;
            }
            string compatibility = element.Normalize(System.Text.NormalizationForm.FormKC);
            string replacement = !string.Equals(compatibility, element, StringComparison.Ordinal) &&
                IsDrawingTextCovered(compatibility, familyName, bold, italic)
                ? compatibility
                : IsDrawingTextCovered(MissingGlyphMark, familyName, bold, italic) ? MissingGlyphMark : "?";
            builder.Append(replacement);
            onSubstitution?.Invoke(element, replacement);
        }
        return builder.ToString();
    }

    private const string MissingGlyphMark = "□";

    private bool HasUncoveredDrawingText(OfficeDrawing drawing, CancellationToken cancellationToken) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (element) {
                case OfficeDrawingText text when PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(text.Text):
                    if (!IsDrawingTextCovered(text.Text, text.Font.FamilyName, text.Font.IsBold, text.Font.IsItalic)) return true;
                    break;
                case OfficeDrawingRichText richText:
                    foreach (OfficeRichTextRun run in richText.Runs) {
                        if (PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(run.Text) &&
                            !IsDrawingTextCovered(run.Text, run.FontFamily, run.Bold, run.Italic)) return true;
                    }
                    break;
                case OfficeDrawingGroup group:
                    if (HasUncoveredDrawingText(group.InnerDrawing, cancellationToken)) return true;
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (HasUncoveredDrawingText(effectGroup.InnerDrawing, cancellationToken)) return true;
                    break;
            }
        }
        return false;
    }

    // Uses the same generated-text analysis as the save preflight, for one run in the drawing text's font.
    private bool IsDrawingTextCovered(string text, string? familyName, bool bold, bool italic) {
        string key = (familyName ?? string.Empty) + "\u0000" + (bold ? "b" : "") + (italic ? "i" : "") + "\u0000" + text;
        _drawingTextCoverage ??= new Dictionary<string, bool>(StringComparer.Ordinal);
        if (_drawingTextCoverage.TryGetValue(key, out bool covered)) return covered;
        PdfStandardFont font = PdfStandardFontMapper.GetFontFamily(DefaultFont);
        if (!string.IsNullOrWhiteSpace(familyName) && PdfStandardFontMapper.TryMapFontFamily(familyName!, out PdfStandardFont mapped)) {
            font = PdfStandardFontMapper.GetFontFamily(mapped);
        }
        var run = new PdfTextRun(text, bold: bold, italic: italic, standardFont: font,
            fontFamily: string.IsNullOrWhiteSpace(familyName) ? null : familyName);
        covered = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(new[] { run }, this, font, "PdfDrawingText").Count == 0;
        _drawingTextCoverage[key] = covered;
        return covered;
    }

    private static void CollectDrawingText(
        OfficeDrawing drawing,
        HashSet<string> families,
        System.Text.StringBuilder unicodeText,
        HashSet<string> seenElements,
        CancellationToken cancellationToken) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (element) {
                case OfficeDrawingText text:
                    AddDrawingText(text.Text, text.Font.FamilyName, families, unicodeText, seenElements);
                    break;
                case OfficeDrawingRichText richText:
                    foreach (OfficeRichTextRun run in richText.Runs) {
                        AddDrawingText(run.Text, run.FontFamily, families, unicodeText, seenElements);
                    }
                    break;
                case OfficeDrawingGroup group:
                    CollectDrawingText(group.InnerDrawing, families, unicodeText, seenElements, cancellationToken);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    CollectDrawingText(effectGroup.InnerDrawing, families, unicodeText, seenElements, cancellationToken);
                    break;
            }
        }
    }

    private static void AddDrawingText(
        string? text,
        string? familyName,
        HashSet<string> families,
        System.Text.StringBuilder unicodeText,
        HashSet<string> seenElements) {
        if (string.IsNullOrEmpty(text)) return;
        if (!string.IsNullOrWhiteSpace(familyName)) families.Add(familyName!.Trim());
        if (!PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(text)) return;
        foreach (string textElement in OfficeTextElements.Enumerate(text!)) {
            if (PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(textElement) && seenElements.Add(textElement)) {
                unicodeText.Append(textElement);
            }
        }
    }
}
