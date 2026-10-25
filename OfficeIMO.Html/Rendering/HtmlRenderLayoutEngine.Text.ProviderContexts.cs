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
                HtmlInlineRun clone = CloneProviderSourceSlice(run, slice.Offset, slice.Length, slice.Shaping?.Text);
                clone.PreparedShaping = slice.Shaping;
                clone.PreparedShapingStyle = slice.Style;
                clone.PreserveContextualFallback = slice.Shaping == null;
                result.Add(clone);
                offset = slice.Offset + slice.Length;
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
        void RetainSlices(InlineWordSlice[] pieces, IReadOnlyList<PreparedInlineSlice> slices) {
            for (int index = 0; index < pieces.Length; index++) {
                CheckCancellation();
                HtmlInlineRun run = pieces[index].Run;
                if (!prepared.TryGetValue(run, out List<PreparedInlineSlice>? retained)) {
                    retained = new List<PreparedInlineSlice>(); prepared.Add(run, retained);
                }
                retained.Add(slices[index]);
            }
        }
        if (!HasArabicJoiningSource(runs, static run => run.LogicalText)) return prepared;
        IOfficeTextShapingProvider provider = _options.TextShapingProvider ?? OfficeManagedTextShapingProvider.Instance;
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
            // Once this complete word owns the provider attempt, a declined or
            // unsplittable result must not turn into partial fragment success at
            // measurement or paint. Preserve the scalar owner's existing forms
            // on these exact source ranges; other words in a run stay eligible.
            void RetainFallback() {
                var fallback = new List<PreparedInlineSlice>(pieces.Length);
                foreach (InlineWordSlice piece in pieces) {
                    CheckCancellation();
                    fallback.Add(new PreparedInlineSlice(piece.Offset, piece.Length, null, Style(piece.Run)));
                }
                RetainSlices(pieces, fallback);
            }
            OfficeTextShapingResult? shaped = ShapeConfiguredText(text, font, fontInfo, style,
                OfficeTextDirection.RightToLeft, provider);
            CheckCancellation();
            if (shaped != null) ChargeLayoutOperations(1L + (shaped.Glyphs.Count + 31L) / 32L, "provider joining glyph retention");
            if (shaped == null) { Report(pieces[0].Run, "cross-inline-provider-declined"); RetainFallback(); continue; }
            if (font is IOfficeColorFontProgram color && shaped.Glyphs.Any(glyph => color.HasColorGlyph(glyph.GlyphId))) {
                Report(pieces[0].Run, "cross-inline-color-glyph-context-not-retained");
                RetainFallback();
                continue;
            }
            var context = new OfficePreparedTextShaping(font, text, shaped, _cancellationToken);
            var slices = new List<PreparedInlineSlice>(pieces.Length);
            int start = 0;
            foreach (InlineWordSlice piece in pieces) {
                CheckCancellation();
                if (!context.TrySlice(start, piece.Length, out OfficePreparedTextShaping? slice)) break;
                slices.Add(new PreparedInlineSlice(piece.Offset, piece.Length, slice!, Style(piece.Run)));
                start += piece.Length;
            }
            if (slices.Count != pieces.Length) { Report(pieces[0].Run, "cross-inline-glyph-slice-not-retained"); RetainFallback(); continue; }
            double completeAdvance = context.Measure(fontInfo.Size);
            double partitionAdvance = slices.Sum(slice => slice.Shaping!.Measure(fontInfo.Size));
            if (Math.Abs(completeAdvance - partitionAdvance) > Math.Max(1D, completeAdvance) * 1e-9D) {
                Report(pieces[0].Run, "cross-inline-nonadditive-advance");
                RetainFallback();
                continue;
            }
            RetainSlices(pieces, slices);
        }
        return prepared;
    }

    private static bool IsProviderContextBarrier(HtmlInlineRun run) =>
        run.Text.Length != run.LogicalText.Length || run.AtomicBlock != null || run.FloatingBlock != null
        || run.PositionedMarkerElement != null || run.RunningStringElement != null || run.RunningElementAssignment != null
        || run.IsEmptyInlineBox || run.LeaderPattern != null || run.IsSuppressedWhitespaceBreak
        || OfficeTextElements.ContainsBidiControl(run.LogicalText) || run.LogicalText.IndexOf('\u00AD') >= 0
        // ZWNJ authors an independent joining context. Do not shape its preceding
        // grapheme across paint scopes, then diagnose the intentional boundary as
        // an unsplittable provider cluster. ZWJ remains transparent here.
        || run.LogicalText.IndexOf('\u200C') >= 0
        || run.InlineEdgeBoundary != null && HasArabicJoiningInsets(run.Style);

    private static bool CanShareProviderContext(HtmlRenderBoxStyle left, HtmlRenderBoxStyle right) =>
        left.Font.Size > 0D && left.Font.Size == right.Font.Size
        && string.Equals(left.Font.FamilyName, right.Font.FamilyName, StringComparison.OrdinalIgnoreCase)
        && left.Font.Face.Equals(right.Font.Face)
        && (left.Font.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic))
            == (right.Font.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic))
        && left.FontDescriptor == right.FontDescriptor
        && left.TextFeatureSettings.Equals(right.TextFeatureSettings)
        && string.Equals(left.Language, right.Language, StringComparison.OrdinalIgnoreCase)
        && left.Direction == right.Direction && left.UnicodeBidi == "normal" && right.UnicodeBidi == "normal"
        && left.WritingMode == "horizontal-tb" && right.WritingMode == "horizontal-tb"
        && left.LetterSpacing == 0D && right.LetterSpacing == 0D
        && left.WordSpacing.Equals(right.WordSpacing)
        && left.WordBreak == "normal" && right.WordBreak == "normal"
        && left.OverflowWrap == "normal" && right.OverflowWrap == "normal"
        && !HasArabicJoiningInsets(left) && !HasArabicJoiningInsets(right);

    private void ReportUnretainedProviderPaint(HtmlInlineRun run, string text) {
        if (!run.PreserveContextualFallback
            && (_options.TextShapingProvider != null || !run.Style.TextFeatureSettings.IsDefault
                || RequiresConfiguredTextShaping(text, run.Style.TextFeatureSettings))
            && TryMeasureWithConfiguredProvider(text, run.Style, out _))
            ReportProviderContextLoss(run, "directional-glyph-context-not-retained");
    }

    private void ReportProviderContextLoss(HtmlInlineRun run, string detail) {
        if (!TryReportComplexTextShapingSource(run.OwnerElement, run.Source)) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            "A shaping context could not retain complete provider glyphs with their authored paint and direction; separate source runs were used.",
            HtmlDiagnosticSeverity.Warning, run.Source, detail, OfficeConversionLossKind.Approximation);
    }

    // Paragraph bidi owns the direction of neutral slices too. Retain its accepted
    // result rather than asking a later backend to infer direction from punctuation.
    private bool TryPrepareDirectionalProviderText(string text, HtmlRenderBoxStyle style,
        OfficeTextDirection direction, double expectedAdvance, out OfficePreparedTextShaping? prepared) =>
        TryPrepareProviderText(text, style, direction, expectedAdvance, resolvedDirectionalRun: true, out prepared);

    private bool TryPrepareProviderText(string text, HtmlRenderBoxStyle style,
        OfficeTextDirection direction, double expectedAdvance, bool resolvedDirectionalRun,
        out OfficePreparedTextShaping? prepared) {
        prepared = null;
        if (_options.TextShapingProvider == null && style.TextFeatureSettings.IsDefault && !OfficeManagedTextShaper.RequiresComplexLayout(text)
            || style.LetterSpacing != 0D || style.WordSpacing != 0D
            || style.WritingMode != "horizontal-tb") return false;
        string logical = OfficeArabicTextShaper.ToLogicalText(text);
        OfficeFontInfo fontInfo = GetEffectiveTextFont(style);
        if (!_fonts.TryResolveFaceForText(logical, fontInfo.FamilyName, style.FontDescriptor, fontInfo.Size,
                out OfficeFontFace? face) || face?.Program is not IOfficeBoundedFontProgram
            || face.Style != (fontInfo.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic))) return false;
        ChargeLayoutOperations(1L + (logical.Length + 31L) / 32L, "directional glyph retention");
        OfficeTextShapingResult? shaped = ShapeConfiguredText(logical, face.Program, fontInfo, style,
            direction, _options.TextShapingProvider ?? OfficeManagedTextShapingProvider.Instance,
            resolvedDirectionalRun: resolvedDirectionalRun);
        if (shaped == null || face.Program is IOfficeColorFontProgram color
            && shaped.Glyphs.Any(glyph => color.HasColorGlyph(glyph.GlyphId))) return false;
        ChargeLayoutOperations(1L + (shaped.Glyphs.Count + 31L) / 32L, "directional glyph retention");
        var candidate = new OfficePreparedTextShaping(face.Program, logical, shaped, _cancellationToken);
        double measuredAdvance = candidate.Measure(fontInfo.Size);
        // The layout width and retained font pens must describe the same run.
        // A new directional result must not be stretched to conceal a different
        // paragraph-context measurement. The caller retains diagnosed fallback.
        if (Math.Abs(measuredAdvance - expectedAdvance) > Math.Max(1D, Math.Abs(expectedAdvance)) * 1e-9D) return false;
        prepared = candidate;
        return true;
    }

    private OfficeTextShapingResult? ShapeConfiguredText(string logicalText, IOfficeFontProgram font,
        OfficeFontInfo fontInfo, HtmlRenderBoxStyle style, OfficeTextDirection direction,
        IOfficeTextShapingProvider provider, bool resolvedDirectionalRun = false) {
        CheckCancellation();
        OfficeTextShapingResult? result = provider.ShapeText(new OfficeTextShapingRequest(
            logicalText, font.DisplayName ?? fontInfo.FamilyName, font.GetFontDataForShaping(), font.IsOpenTypeCff,
            font.UnitsPerEm, direction, _options.TextShapingLanguage ?? style.Language, _cancellationToken, font.CollectionIndex,
            (font as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping, cloneFontData: false,
            fontProgramCacheKey: font, featureSettings: style.TextFeatureSettings,
            resolvedDirectionalRun: resolvedDirectionalRun));
        CheckCancellation();
        return result;
    }

    private readonly struct PreparedInlineSlice {
        internal PreparedInlineSlice(int offset, int length, OfficePreparedTextShaping? shaping, HtmlRenderBoxStyle style) {
            Offset = offset; Length = length; Shaping = shaping; Style = style;
        }
        internal int Offset { get; }
        internal int Length { get; }
        internal OfficePreparedTextShaping? Shaping { get; }
        internal HtmlRenderBoxStyle Style { get; }
    }
}
