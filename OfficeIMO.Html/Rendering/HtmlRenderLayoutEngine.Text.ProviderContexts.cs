using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // A retained result belongs to one source word and font instance. Paint-only
    // boundaries may divide complete clusters, but never trigger new shaping.
    private List<HtmlInlineRun> PrepareConfiguredJoiningRuns(IReadOnlyList<HtmlInlineRun> runs,
        HtmlRenderBoxStyle paragraphStyle, Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector = null, bool reportLoss = true) {
        Dictionary<HtmlInlineRun, List<PreparedInlineSlice>> prepared = PrepareConfiguredJoiningSlices(runs, paragraphStyle, styleSelector, reportLoss);
        var result = new List<HtmlInlineRun>(runs.Count);
        foreach (HtmlInlineRun run in runs) {
            if (!prepared.TryGetValue(run, out List<PreparedInlineSlice>? slices)) { result.Add(run); continue; }
            int offset = 0;
            foreach (PreparedInlineSlice slice in slices) {
                CheckCancellation();
                if (slice.Offset > offset) result.Add(CloneProviderSourceSlice(run, offset, slice.Offset - offset));
                HtmlInlineRun clone = CloneProviderSourceSlice(run, slice.Offset, slice.Shaping.Text.Length, slice.Shaping.Text);
                clone.PreparedShaping = slice.Shaping;
                clone.PreparedShapingStyle = slice.Style;
                result.Add(clone);
                offset = slice.Offset + slice.Shaping.Text.Length;
            }
            if (offset < run.Text.Length) result.Add(CloneProviderSourceSlice(run, offset, run.Text.Length - offset));
        }
        return result;
    }

    private static HtmlInlineRun CloneProviderSourceSlice(HtmlInlineRun run, int offset, int length, string? preparedText = null) {
        HtmlInlineRun clone = run.CloneText(preparedText ?? run.Text.Substring(offset, length),
            run.LogicalText.Substring(offset, length), run.Style, run.IsFirstLetter);
        clone.PreparedHyphenation = run.PreparedHyphenation?.SliceSource(offset, length);
        clone.EndsFirstLine = run.EndsFirstLine && offset + length == run.Text.Length;
        clone.FirstLineHyphen = clone.EndsFirstLine ? run.FirstLineHyphen : string.Empty;
        return clone;
    }

    private Dictionary<HtmlInlineRun, List<PreparedInlineSlice>> PrepareConfiguredJoiningSlices(
        IReadOnlyList<HtmlInlineRun> runs, HtmlRenderBoxStyle paragraphStyle,
        Func<HtmlInlineRun, HtmlRenderBoxStyle>? styleSelector = null, bool reportLoss = true) {
        var prepared = new Dictionary<HtmlInlineRun, List<PreparedInlineSlice>>();
        void Report(HtmlInlineRun run, string detail) { if (reportLoss) ReportProviderContextLoss(run, detail); }
        if (_options.TextShapingProvider == null || !HasArabicJoiningSource(runs, static run => run.LogicalText)) return prepared;
        InlineWordMap map = BuildInlineWordFitting(runs, paragraphStyle, styleSelector);
        var styles = new Dictionary<HtmlInlineRun, HtmlRenderBoxStyle>();
        HtmlRenderBoxStyle Style(HtmlInlineRun run) {
            if (!styles.TryGetValue(run, out HtmlRenderBoxStyle? style)) {
                style = styleSelector?.Invoke(run) ?? run.Style;
                styles.Add(run, style);
            }
            return style;
        }
        foreach (InlineWordSlice[] word in map.CrossRunWords()) {
            CheckCancellation();
            InlineWordSlice[] pieces = word.Where(static piece => piece.Length > 0).ToArray();
            HtmlRenderBoxStyle style = Style(pieces[0].Run);
            if (word.Any(piece => IsProviderContextBarrier(piece.Run))
                || pieces.Any(piece => !CanShareProviderContext(style, Style(piece.Run)))) continue;
            var source = new StringBuilder();
            foreach (InlineWordSlice piece in pieces) source.Append(piece.Run.LogicalText, piece.Offset, piece.Length);
            string text = source.ToString();
            if (!OfficeTextElements.ContainsJoiningScript(text)) continue;
            ChargeLayoutOperations(1L + (text.Length + 31L) / 32L, "provider joining context");
            IReadOnlyList<OfficeBidiTextRun> directional = OfficeBidiTextResolver.ResolveVisualRuns(text, OfficeTextDirection.RightToLeft);
            if (directional.Count != 1 || directional[0].Direction != OfficeTextDirection.RightToLeft) continue;
            OfficeFontInfo fontInfo = GetEffectiveTextFont(style);
            if (!_fonts.TryResolveFaceForText(text, fontInfo.FamilyName, style.FontDescriptor, fontInfo.Size, out OfficeFontFace? face)
                || face == null || face.Program is not IOfficeBoundedFontProgram
                || face.Style != (fontInfo.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic))) {
                Report(pieces[0].Run, "cross-inline-font-context-not-retained");
                continue;
            }
            IOfficeFontProgram font = face.Program;
            bool sameProgram = true;
            foreach (InlineWordSlice piece in pieces) {
                CheckCancellation();
                if (!_fonts.TryResolveFaceForText(piece.Run.LogicalText.Substring(piece.Offset, piece.Length),
                    fontInfo.FamilyName, style.FontDescriptor, fontInfo.Size, out OfficeFontFace? selected)
                    || !ReferenceEquals(selected?.Program, font)) { sameProgram = false; break; }
            }
            if (!sameProgram) { Report(pieces[0].Run, "cross-inline-font-context-not-retained"); continue; }
            OfficeTextShapingResult? shaped = _options.TextShapingProvider.ShapeText(new OfficeTextShapingRequest(
                text, font.DisplayName ?? fontInfo.FamilyName, font.GetFontDataForShaping(), font.IsOpenTypeCff, font.UnitsPerEm,
                OfficeTextDirection.RightToLeft, _options.TextShapingLanguage, _cancellationToken, font.CollectionIndex,
                (font as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping, cloneFontData: false,
                fontProgramCacheKey: font, featureSettings: style.TextFeatureSettings));
            CheckCancellation();
            if (shaped != null) ChargeLayoutOperations(1L + (shaped.Glyphs.Count + 31L) / 32L, "provider joining glyph retention");
            if (shaped == null) { Report(pieces[0].Run, "cross-inline-provider-declined"); continue; }
            if (font is IOfficeColorFontProgram color && shaped.Glyphs.Any(glyph => color.HasColorGlyph(glyph.GlyphId))) {
                Report(pieces[0].Run, "cross-inline-color-glyph-context-not-retained");
                continue;
            }
            var context = new OfficePreparedTextShaping(font, text, shaped, _cancellationToken);
            var slices = new List<PreparedInlineSlice>(pieces.Length);
            int start = 0;
            foreach (InlineWordSlice piece in pieces) {
                CheckCancellation();
                if (!context.TrySlice(start, piece.Length, out OfficePreparedTextShaping? slice)) break;
                slices.Add(new PreparedInlineSlice(piece.Offset, slice!, Style(piece.Run)));
                start += piece.Length;
            }
            if (slices.Count != pieces.Length) { Report(pieces[0].Run, "cross-inline-glyph-slice-not-retained"); continue; }
            double completeAdvance = context.Measure(fontInfo.Size);
            double partitionAdvance = slices.Sum(slice => slice.Shaping.Measure(fontInfo.Size));
            if (Math.Abs(completeAdvance - partitionAdvance) > Math.Max(1D, completeAdvance) * 1e-9D) {
                Report(pieces[0].Run, "cross-inline-nonadditive-advance");
                continue;
            }
            for (int index = 0; index < pieces.Length; index++) {
                HtmlInlineRun run = pieces[index].Run;
                if (!prepared.TryGetValue(run, out List<PreparedInlineSlice>? retained)) {
                    retained = new List<PreparedInlineSlice>(); prepared.Add(run, retained);
                }
                retained.Add(slices[index]);
            }
        }
        return prepared;
    }

    private static bool IsProviderContextBarrier(HtmlInlineRun run) =>
        run.Text.Length != run.LogicalText.Length || run.AtomicBlock != null || run.FloatingBlock != null
        || run.PositionedMarkerElement != null || run.RunningStringElement != null || run.RunningElementAssignment != null
        || run.IsEmptyInlineBox || run.LeaderPattern != null || run.IsSuppressedWhitespaceBreak
        || OfficeTextElements.ContainsBidiControl(run.LogicalText) || run.LogicalText.IndexOf('\u00AD') >= 0
        || run.InlineEdgeBoundary != null && HasArabicJoiningInsets(run.Style);

    private static bool CanShareProviderContext(HtmlRenderBoxStyle left, HtmlRenderBoxStyle right) =>
        left.Font.Size > 0D && left.Font.Equals(right.Font) && left.FontDescriptor == right.FontDescriptor
        && left.TextFeatureSettings.Equals(right.TextFeatureSettings)
        && string.Equals(left.Language, right.Language, StringComparison.OrdinalIgnoreCase)
        && left.Direction == right.Direction && left.UnicodeBidi == "normal" && right.UnicodeBidi == "normal"
        && left.WritingMode == "horizontal-tb" && right.WritingMode == "horizontal-tb"
        && left.LetterSpacing == 0D && right.LetterSpacing == 0D
        && left.WordSpacing.Equals(right.WordSpacing)
        && left.WordBreak == "normal" && right.WordBreak == "normal"
        && left.OverflowWrap == "normal" && right.OverflowWrap == "normal"
        && !HasArabicJoiningInsets(left) && !HasArabicJoiningInsets(right);

    private void ReportProviderContextLoss(HtmlInlineRun run, string detail) {
        if (run.OwnerElement != null && !_reportedComplexTextShapingElements.Add(run.OwnerElement)) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            "A cross-inline shaping context could not retain complete provider glyphs with their authored paint; separate source runs were used.",
            HtmlDiagnosticSeverity.Warning, run.Source, detail, OfficeConversionLossKind.Approximation);
    }

    private readonly struct PreparedInlineSlice {
        internal PreparedInlineSlice(int offset, OfficePreparedTextShaping shaping, HtmlRenderBoxStyle style) {
            Offset = offset; Shaping = shaping; Style = style;
        }
        internal int Offset { get; }
        internal OfficePreparedTextShaping Shaping { get; }
        internal HtmlRenderBoxStyle Style { get; }
    }
}
