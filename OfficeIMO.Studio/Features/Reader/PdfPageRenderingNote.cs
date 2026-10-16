using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>What a page's rendering notes mean for the reader, from least to most significant.</summary>
public enum PdfPageRenderingNote {
    /// <summary>No notes, or only notes about which renderer drew the page.</summary>
    None,
    /// <summary>A standard PDF font the document does not embed is drawn with an installed font, as every viewer does.</summary>
    SubstituteFont,
    /// <summary>Some details are drawn in a documented, simplified way.</summary>
    Simplified,
    /// <summary>A font program the document relies on could not be used, so letter shapes can differ.</summary>
    FontReplaced,
    /// <summary>Some content is skipped because the renderer does not support it.</summary>
    ContentOmitted
}

/// <summary>
/// Classifies renderer diagnostics through the engine's capability manifest, so Studio decides from support levels and
/// the engine's own font distinction rather than from engine wording.
/// </summary>
internal static class PdfPageRenderingNotes {
    private const string FontSubstitution = "render.resource.font-substitution";
    private static readonly HashSet<string> EmbeddedFontSubstitutions = new(StringComparer.Ordinal) {
        "render.resource.font-cff-substitution",
        "render.resource.font-type3-substitution"
    };
    // Annotations without an appearance stream (most often links) usually draw nothing in other viewers either, so
    // they are a page note rather than a reason to doubt the page.
    private const string AnnotationAppearanceMissing = "render.resource.annotation-appearance-missing";
    private static readonly Dictionary<string, PdfRenderSupportLevel> SupportLevels =
        PdfRenderCapabilities.Current.Entries.ToDictionary(entry => entry.Id, entry => entry.SupportLevel, StringComparer.Ordinal);

    /// <summary>Classifies the page's typed capability diagnostics and the page renderer's "code: message" notes.</summary>
    internal static PdfPageRenderingNote Classify(IEnumerable<PdfRenderCapabilityDiagnostic> capabilities, IEnumerable<string> rendered) {
        PdfPageRenderingNote note = Classify(rendered);
        foreach (PdfRenderCapabilityDiagnostic capability in capabilities) {
            PdfPageRenderingNote current = capability.Code == FontSubstitution && !capability.IsStandardFontSubstitution
                ? PdfPageRenderingNote.FontReplaced
                : Classify(capability.Code, capability.SupportLevel);
            if (current > note) note = current;
        }
        return note;
    }

    /// <summary>
    /// Classifies codes or "code: message" notes. A bare font substitution is read as the routine standard-font case; the
    /// typed overload tells the two apart.
    /// </summary>
    internal static PdfPageRenderingNote Classify(IEnumerable<string> codesOrDiagnostics) {
        var note = PdfPageRenderingNote.None;
        foreach (string entry in codesOrDiagnostics) {
            int separator = entry.IndexOf(':');
            string code = separator > 0 ? entry[..separator] : entry;
            if (!SupportLevels.TryGetValue(code, out PdfRenderSupportLevel level)) continue;
            PdfPageRenderingNote current = Classify(code, level);
            if (current > note) note = current;
        }
        return note;
    }

    private static PdfPageRenderingNote Classify(string code, PdfRenderSupportLevel level) => level switch {
        _ when code == AnnotationAppearanceMissing => PdfPageRenderingNote.Simplified,
        PdfRenderSupportLevel.Unsupported => PdfPageRenderingNote.ContentOmitted,
        _ when EmbeddedFontSubstitutions.Contains(code) => PdfPageRenderingNote.FontReplaced,
        _ when code == FontSubstitution => PdfPageRenderingNote.SubstituteFont,
        PdfRenderSupportLevel.Simplified => PdfPageRenderingNote.Simplified,
        _ => PdfPageRenderingNote.None
    };

    /// <summary>Notes worth a document-level notice: the page can look noticeably different from the original.</summary>
    internal static bool AffectsAppearance(PdfPageRenderingNote note) => note >= PdfPageRenderingNote.FontReplaced;

    internal static string? ExplanationKey(PdfPageRenderingNote note) => note switch {
        PdfPageRenderingNote.SubstituteFont => "PdfPage.FontSubstitution",
        PdfPageRenderingNote.Simplified => "PdfPage.RenderingSimplified",
        PdfPageRenderingNote.FontReplaced => "PdfPage.FontReplaced",
        PdfPageRenderingNote.ContentOmitted => "PdfPage.RenderingLimitation",
        _ => null
    };
}
