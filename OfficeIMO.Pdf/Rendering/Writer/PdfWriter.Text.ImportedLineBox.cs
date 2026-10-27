namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Resolves a lower bound for imported paragraph layout from the same selected
    /// font programs and fallback runs as painting. Actual line heights remain
    /// owned by the rich-line layout, including mixed-font ascent and descent.
    /// </summary>
    /// <remarks>
    /// Fallback normalization records font-resource ownership on <paramref name="options"/>.
    /// A missing font program or inherited body font size leaves the caller's
    /// generic spacing policy intact. The latter must not be rebound by changing
    /// the paragraph's layout size beneath a run with no explicit size.
    /// The minimum, rather than maximum, ratio prevents a font on another line
    /// from imposing a paragraph-wide height on otherwise smaller text.
    /// </remarks>
    internal static bool TryResolveImportedLineBoxFloor(
        IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont,
        PdfCreateOptions options, out double layoutFontSize, out double naturalMultiplier) {
        layoutFontSize = double.PositiveInfinity;
        naturalMultiplier = double.PositiveInfinity;
        foreach (PdfTextRun run in NormalizeFallbackRuns(runs, baseFont, options)) {
            if (!run.ContributesLineBoxDescent || run.InlineElement != null || IsLayoutControlRun(run) ||
                string.IsNullOrEmpty(run.Text) && !run.IsLineBoxStrut) continue;
            if (run.FontSize is not double nominalSize || nominalSize <= 0D ||
                double.IsNaN(nominalSize) || double.IsInfinity(nominalSize)) return false;
            PdfStandardFont font = ResolveFontForRun(run, baseFont);
            PdfNamedFontFace? named = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic,
                out PdfNamedFontFace resolved) ? resolved : null;
            var metrics = ResolveFontLineMetrics(font, named, options);
            if (!metrics.HasValue) return false;

            double effectiveSize = EffectiveRichFontSize(nominalSize, run.Baseline);
            layoutFontSize = Math.Min(layoutFontSize, nominalSize);
            // RunLineHeight receives the nominal size. Normalize the natural
            // advance by that same size so raised/lowered text cannot create
            // a larger unscaled body strut before per-run layout takes over.
            naturalMultiplier = Math.Min(naturalMultiplier,
                metrics.Value.HorizontalAdvanceRatio * effectiveSize / nominalSize);
        }
        return !double.IsInfinity(layoutFontSize) && !double.IsInfinity(naturalMultiplier);
    }
}
