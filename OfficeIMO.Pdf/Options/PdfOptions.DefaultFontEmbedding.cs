using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private static readonly PdfStandardFont[] DefaultFontFamilySlots = {
        PdfStandardFont.Helvetica, PdfStandardFont.TimesRoman, PdfStandardFont.Courier
    };
    private bool _embedDefaultFonts = EmbedDefaultFontsInitially;
    private HashSet<PdfStandardFont>? _defaultFontEmbeddingAttempts;
    private HashSet<PdfStandardFont>? _defaultEmbeddedFontFamilies;
    private bool _defaultFontsUseStandardMetrics;

    // Test assemblies that pin standard-font metrics turn this off once at start-up; products never change it.
    internal static bool EmbedDefaultFontsInitially { get; set; } = true;

    /// <summary>
    /// When true (the default), text set in a standard font slot (Helvetica, Times or Courier) that has no registered
    /// font file uses an installed font resolved through the default theme's fallback lists, embedded as a subset:
    /// the body font for Helvetica, the serif chain for Times and the monospace chain for Courier. When no font file
    /// can be found the standard font is written without embedding and a <c>font-not-embedded</c> diagnostic is reported.
    /// Set to false to write the non-embedded standard PDF fonts.
    /// </summary>
    public bool EmbedDefaultFonts {
        get => _embedDefaultFonts;
        set => _embedDefaultFonts = value;
    }

    /// <summary>
    /// Backs every standard font slot that has no registered font file with an installed font from the theme
    /// fallback lists. Converters call this after configuring source fonts and before measuring text, so layout
    /// uses the metrics of the font that is embedded; the PDF writer calls it before layout for native documents.
    /// </summary>
    /// <param name="standardMetrics">
    /// True for converters: a slot they leave unregistered stands for a source family such as Arial, so it takes a
    /// metric-compatible face first. Native documents use the theme body font for the default sans-serif slot.
    /// </param>
    internal void ApplyDefaultFontEmbedding(bool standardMetrics = false) {
        if (!_embedDefaultFonts) return;
        _defaultFontsUseStandardMetrics |= standardMetrics;
        foreach (PdfStandardFont slot in DefaultFontFamilySlots) {
            if (HasAnyEmbeddedStyle(slot)) continue;
            // A fallback set claims its listed slots and, when those are occupied, any free slot at render time.
            // Only the body slot is backed then, so fallback runs keep somewhere to go.
            if (_embeddedFontFallbacks != null &&
                (slot != PdfStandardFontMapper.GetFontFamily(DefaultFont) ||
                 _embeddedFontFallbacks.FontSlots.Any(fallback => PdfStandardFontMapper.GetFontFamily(fallback) == slot))) continue;
            if (!(_defaultFontEmbeddingAttempts ??= new HashSet<PdfStandardFont>()).Add(slot)) continue;
            if (!TryResolveDefaultFontFamily(slot, standardMetrics, out PdfEmbeddedFontFamily? family) || family == null) continue;
            RegisterFontFamily(slot, family);
            (_defaultEmbeddedFontFamilies ??= new HashSet<PdfStandardFont>()).Add(slot);
        }
    }

    /// <summary>Reports standard fonts that the generated document uses without an embedded font file.</summary>
    internal void ReportNonEmbeddedStandardFonts(IEnumerable<PdfGeneratedFontComplianceEvidence> usages) {
        if (DiagnosticsReport == null) return;
        var reported = new HashSet<PdfStandardFont>();
        foreach (PdfGeneratedFontComplianceEvidence usage in usages) {
            if (!usage.StandardFont.HasValue) continue;
            PdfStandardFont font = usage.StandardFont.Value;
            PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(font);
            if (usage.Options.TryGetEmbeddedStandardFont(font, out PdfEmbeddedFont? embedded) && embedded != null) continue;
            if (!reported.Add(family)) continue;
            // Turning embedding off is an explicit request for standard fonts and needs no diagnostic.
            if (!usage.Options._embedDefaultFonts) continue;
            const string reason = "FontUnavailable";
            if (!PdfOutputDiagnosticScope.Admit("font-not-embedded|" + family, ref _reportedLayoutDiagnostics)) continue;
            DiagnosticsReport.Add(new PdfConversionWarning(
                _diagnosticsConverter, "font-not-embedded", "font:" + family.ToBaseFontName(),
                "No installed font from the theme fallback list could be embedded for '" + family.ToBaseFontName() + "'; text uses the non-embedded standard PDF font and its appearance depends on the viewer.",
                // The text itself is intact; only its appearance depends on the viewer, so this is not reported as loss.
                PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.None,
                details: new Dictionary<string, string> {
                    ["standardFont"] = family.ToBaseFontName(),
                    ["reason"] = reason,
                    ["fallbackList"] = string.Join(", ", EnumerateDefaultFontFamilyCandidates(family, usage.Options._defaultFontsUseStandardMetrics))
                }));
        }
    }

    // A library-selected default font was not requested by the source, so leaving its optional contextual
    // alternates unapplied simplifies nothing the author asked for; the source text is set in the font's default forms.
    private bool IsDefaultEmbeddedFontSubstitutionDiagnostic(PdfTextShapingDiagnostic diagnostic, string? fontName) {
        if (_defaultEmbeddedFontFamilies == null || _embeddedFonts == null ||
            !string.Equals(diagnostic.Code, "unsupported-font-ligature-substitution", StringComparison.Ordinal)) return false;
        foreach (PdfStandardFont family in _defaultEmbeddedFontFamilies) {
            for (int style = 0; style < 4; style++) {
                PdfStandardFont font = PdfStandardFontMapper.GetStyledFont(family, bold: (style & 1) != 0, italic: (style & 2) != 0);
                if (!_embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? embedded) || string.IsNullOrEmpty(embedded.FontName)) continue;
                if (string.Equals(fontName, embedded.FontName, StringComparison.Ordinal) ||
                    diagnostic.Message.Contains("'" + embedded.FontName + "'")) return true;
            }
        }
        return false;
    }

    internal bool IsDefaultEmbeddedFontFamily(PdfStandardFont slot) =>
        _defaultEmbeddedFontFamilies?.Contains(PdfStandardFontMapper.GetFontFamily(slot)) == true;

    private bool HasAnyEmbeddedStyle(PdfStandardFont slot) {
        PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(slot);
        for (int style = 0; style < 4; style++) {
            PdfStandardFont font = PdfStandardFontMapper.GetStyledFont(family, bold: (style & 1) != 0, italic: (style & 2) != 0);
            if (_embeddedFonts != null && _embeddedFonts.ContainsKey(font)) return true;
        }
        return false;
    }

    private static bool TryResolveDefaultFontFamily(PdfStandardFont slot, bool standardMetrics, out PdfEmbeddedFontFamily? family) {
        family = null;
        PdfEmbeddedFontFamily? first = null;
        foreach (string candidate in EnumerateDefaultFontFamilyCandidates(slot, standardMetrics)) {
            if (!PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfEmbeddedFontFamily? loaded) || loaded == null) continue;
            // Prefer a family with real bold and italic faces so emphasis survives the fallback.
            if (loaded.BoldSnapshot != null && loaded.ItalicSnapshot != null) {
                family = loaded;
                return true;
            }
            first ??= loaded;
        }
        family = first;
        return family != null;
    }

    internal static IEnumerable<string> EnumerateDefaultFontFamilyCandidates(PdfStandardFont slot, bool standardMetrics = false) {
        PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(slot);
        if (family == PdfStandardFont.Helvetica && !standardMetrics) {
            OfficeThemeFont body = OfficeTheme.Default.Fonts.Body;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (seen.Add(body.Family)) yield return body.Family;
            foreach (string fallback in body.Fallbacks) {
                if (!IsGenericFamilyKeyword(fallback) && seen.Add(fallback)) yield return fallback;
            }
            foreach (string substitute in OfficeDocumentFontSubstitutions.EnumerateGeneric("sans-serif")) {
                if (seen.Add(substitute)) yield return substitute;
            }
            yield break;
        }
        // Native code text takes the modern monospace chain (Consolas first); converters keep Courier metrics.
        string standard = family == PdfStandardFont.Courier ? (standardMetrics ? "Courier" : "monospace")
            : family == PdfStandardFont.TimesRoman ? "Times" : "Helvetica";
        foreach (string substitute in OfficeDocumentFontSubstitutions.EnumerateGeneric(standard)) yield return substitute;
    }

    private static bool IsGenericFamilyKeyword(string family) {
        string key = family.Trim().ToLowerInvariant();
        return key == "sans-serif" || key == "serif" || key == "monospace" || key == "cursive" || key == "fantasy" || key == "system-ui";
    }
}
