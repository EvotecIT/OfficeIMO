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
