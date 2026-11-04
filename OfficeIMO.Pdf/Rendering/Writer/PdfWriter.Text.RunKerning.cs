using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Position compatible glyph spans through Core before distributing them back
    // to their paint owners. GPOS lookup consumption belongs to the whole span.
    private sealed class RichRunKerning {
        private const long MaximumKerningWork = 1_000_000;
        private long _work;
        private readonly PdfCreateOptions? _options;
        private Dictionary<RichSeg, PdfTextShowCommand>? _commands;
        private Dictionary<(PdfStandardFont, PdfNamedFontFace?), RunKerningFont?>? _fonts;
        private Dictionary<(RichSeg, RichSeg, int), double>? _adjustments;

        public RichRunKerning(PdfCreateOptions? options) => _options = options;

        private void SpendWork(long amount) {
            // Charge before scanning/copying a compatible span or positioning its glyphs.
            // Repeated whole-prefix measurements must remain bounded even on a very wide page.
            if (amount > MaximumKerningWork - _work)
                throw new System.IO.InvalidDataException("PDF text layout exceeds the cross-run kerning work limit.");
            _work += amount;
        }

        public double MeasureBoundary(IReadOnlyList<RichSeg> line, RichSeg right) {
            if (line.Count == 0 || !Compatible(line[line.Count - 1], right)) return 0D;
            int start = line.Count - 1;
            while (start > 0 && Compatible(line[start - 1], line[start])) start--;
            SpendWork(line.Count - start);
            var tail = new RichSeg[line.Count - start];
            for (int index = 0; index < tail.Length; index++) tail[index] = line[start + index];
            return MeasureJoin(tail, new[] { right });
        }

        public double MeasureJoin(IReadOnlyList<RichSeg>? left, IReadOnlyList<RichSeg>? right) {
            if (left == null || right == null || left.Count == 0 || right.Count == 0
                || !Compatible(left[left.Count - 1], right[0])) return 0D;
            SpendWork((long)left.Count + right.Count);
            return GetAdjustment(left.Concat(right).ToArray()) - GetAdjustment(left) - GetAdjustment(right);
        }

        public void Apply(IReadOnlyList<RichSeg> line) => PositionSpans(line, apply: true);

        public double MeasureWordAdjustment(RichSeg first, IReadOnlyList<RichSeg>? remaining) {
            if (remaining == null || remaining.Count == 0) return 0D;
            if (!HasKerningFeature(first) && !remaining.Any(HasKerningFeature)) return 0D;
            SpendWork((long)remaining.Count + 1);
            var segments = new RichSeg[remaining.Count + 1];
            segments[0] = first;
            for (int index = 0; index < remaining.Count; index++) segments[index + 1] = remaining[index];
            return GetAdjustment(segments);
        }

        private double GetAdjustment(IReadOnlyList<RichSeg> segments) {
            if (segments.Count < 2) return 0D;
            if (!segments.Any(HasKerningFeature)) return 0D;
            _adjustments ??= new();
            var key = (segments[0], segments[segments.Count - 1], segments.Count);
            if (!_adjustments.TryGetValue(key, out double adjustment)) {
                adjustment = PositionSpans(segments, apply: false);
                _adjustments.Add(key, adjustment);
            }
            return adjustment;
        }

        private static bool HasKerningFeature(RichSeg segment) =>
            segment.FeatureSettings.TryGetValue("kern", out int kern) && kern > 0;

        private bool Compatible(RichSeg left, RichSeg right) {
            bool compatible = _options != null
            && left.InlineElement == null && right.InlineElement == null
            && !left.PreservesPaintedGlyphs && !right.PreservesPaintedGlyphs
            && !left.EndsWithHardBreak && !left.EndsWithTextSeparator && !right.LeadingSpace && !right.LeadingIsTab
            && left.Text.Length > 0 && right.Text.Length > 0
            && !char.IsWhiteSpace(left.Text, left.Text.Length - 1) && !char.IsWhiteSpace(right.Text, 0)
            && !char.IsControl(left.Text, left.Text.Length - 1) && !char.IsControl(right.Text, 0)
            && left.Font == right.Font && left.NamedFont == right.NamedFont
            && left.FontSize == right.FontSize && left.Baseline == right.Baseline
            && left.HorizontalTextScaling == right.HorizontalTextScaling && left.CharacterSpacing == right.CharacterSpacing
            && left.FontMetricScale == 1D && right.FontMetricScale == 1D
            && left.FeatureSettings.Equals(right.FeatureSettings)
            && left.FeatureSettings.TryGetValue("kern", out int kern) && kern > 0
            && left.TextDirection is OfficeTextDirection.Auto or OfficeTextDirection.LeftToRight
            && right.TextDirection is OfficeTextDirection.Auto or OfficeTextDirection.LeftToRight
            && (_options.TextShapingProviderSnapshot == null || ReferenceEquals(_options.TextShapingProviderSnapshot, OfficeManagedTextShapingProvider.Instance));
            if (!compatible) return false;
            SpendWork((long)left.Text.Length + right.Text.Length);
            return !OfficeManagedTextShaper.RequiresComplexLayout(left.Text) && !OfficeManagedTextShaper.RequiresComplexLayout(right.Text);
        }

        private double PositionSpans(IReadOnlyList<RichSeg> segments, bool apply) {
            double adjustment = 0D;
            int start = 0;
            while (start < segments.Count) {
                int end = start + 1;
                while (end < segments.Count && Compatible(segments[end - 1], segments[end])) end++;
                if (end - start > 1) adjustment += PositionSpan(segments, start, end, apply);
                start = end;
            }
            return adjustment;
        }

        private double PositionSpan(IReadOnlyList<RichSeg> segments, int start, int end, bool apply) {
            RichSeg first = segments[start];
            RunKerningFont? font = GetFont(first);
            if (font == null) return 0D;
            SpendWork(end - start);
            var commands = new PdfTextShowCommand[end - start];
            var glyphIds = new List<int>();
            var scalars = new List<int>();
            for (int index = start; index < end; index++) {
                PdfTextShowCommand command = GetCommand(segments[index]);
                if (command.VisualGlyphs == null) return 0D;
                commands[index - start] = command;
                string sourceText = WithoutOptionalBreaks(segments[index].Text);
                SpendWork(command.VisualGlyphs.Count);
                foreach (PdfGlyphInfo glyph in command.VisualGlyphs) {
                    glyphIds.Add(glyph.GlyphId);
                    // Multiple substitutions retain the source character even when
                    // their continuation glyphs have no separate extraction text.
                    int sourceIndex = glyph.TextIndex;
                    scalars.Add((uint)sourceIndex < (uint)sourceText.Length
                        ? ReadScalar(sourceText, ref sourceIndex) : glyph.UnicodeScalar);
                }
            }
            SpendWork(glyphIds.Count);
            OfficeOpenTypeGlyphPositioning[] positions = font.Kerning.PositionRun(glyphIds, scalars);
            double adjustment = 0D;
            int glyphIndex = 0;
            for (int index = start; index < end; index++) {
                PdfTextShowCommand command = commands[index - start];
                SpendWork(command.VisualGlyphs!.Count);
                PdfGlyphInfo[] glyphs = command.VisualGlyphs!.ToArray();
                int advanceDelta = 0;
                double fontUnitAdvance = 0D;
                bool changed = false;
                for (int local = 0; local < glyphs.Length; local++, glyphIndex++) {
                    PdfGlyphInfo glyph = glyphs[local];
                    OfficeOpenTypeGlyphPositioning position = positions[glyphIndex];
                    // The managed shaper scales nominal advances and positioning
                    // independently, including substitution continuation glyphs.
                    int advance = checked(glyph.NominalWidth1000 + font.Scale(position.XAdvance));
                    int x = font.Scale(position.XPlacement);
                    if (_options?.UseFontUnitTextLayout == true)
                        fontUnitAdvance += PdfFontUnitLayoutMetrics.ScaleAdvance(
                            checked(font.Reader.AdvanceWidth(glyph.GlyphId) + position.XAdvance), first.HorizontalTextScaling);
                    advanceDelta = checked(advanceDelta + advance - glyph.AdvanceWidth1000);
                    changed |= advance != glyph.AdvanceWidth1000 || x != glyph.OffsetX1000;
                    if (apply) glyphs[local] = new PdfGlyphInfo(glyph.GlyphId, glyph.UnicodeText, glyph.TextIndex, glyph.NominalWidth1000,
                        advance, glyph.AdvanceHeight1000, x, glyph.OffsetY1000, glyph.LogicalClusterStart);
                }
                double layoutDelta = command.FontUnitLayoutAdvanceWidth1000 is double originalAdvance
                    ? fontUnitAdvance * 1000D / font.Reader.UnitsPerEm - originalAdvance
                    : advanceDelta * first.HorizontalTextScaling / 100D;
                double points = layoutDelta * EffectiveRichFontSize(first.FontSize, first.Baseline) / 1000D;
                adjustment += points;
                if (apply) segments[index].KerningAdvance = points;
                if (apply && changed) {
                    // Positioned operators use glyph IDs even when the ordinary ASCII
                    // fast path uses distinct character codes. Preserve logical clusters.
                    segments[index].KerningCommand = new PdfTextShowCommand(PdfGlyphRun.ToGlyphHex(glyphs), glyphs, command.ActualText,
                        logicalGlyphs: command.ActualText == null ? glyphs : command.LogicalGlyphs != null ? glyphs : null,
                        advanceWidth1000: command.AdvanceWidth1000 + advanceDelta, wordSpaceCount: command.WordSpaceCount,
                        visualGlyphs: glyphs, tracking: command.Tracking, unitsPerEm: command.UnitsPerEm,
                        trackingBoundaries: command.TrackingBoundaries, negativeTracking: command.NegativeTracking,
                        fontMetricScale: command.FontMetricScale, glyphCount: command.GlyphCount, isScalarBidi: command.IsScalarBidi, isolatedRightToLeft: command.IsolatedRightToLeft).WithPinner(command.Pinner);
                }
            }
            return adjustment;
        }

        private PdfTextShowCommand GetCommand(RichSeg segment) {
            _commands ??= new();
            if (!_commands.TryGetValue(segment, out PdfTextShowCommand? command)) {
                command = EncodeTextShowCommand(segment.Text, segment.Font, segment.NamedFont, _options,
                    segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale, segment.HorizontalTextScaling,
                    segment.PreservesPaintedGlyphs);
                _commands.Add(segment, command);
            }
            return command;
        }

        private RunKerningFont? GetFont(RichSeg segment) {
            var key = (segment.Font, segment.NamedFont);
            _fonts ??= new();
            if (_fonts.TryGetValue(key, out RunKerningFont? state)) return state;
            byte[]? data = null;
            if (segment.NamedFont is PdfNamedFontFace named) {
                if (_options!.TryGetNamedFontProgram(named, out PdfTrueTypeFontProgram? font)) data = font!.FontDataForInspection;
                else if (_options.TryGetNamedOpenTypeCffFontProgram(named, out PdfOpenTypeCffFontProgram? cff)) data = cff!.FontDataForInspection;
            } else {
                if (_options!.TryGetEmbeddedStandardFontProgram(segment.Font, out PdfTrueTypeFontProgram? font)) data = font!.FontDataForInspection;
                else if (_options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out PdfOpenTypeCffFontProgram? cff)) data = cff!.FontDataForInspection;
            }
            OfficeOpenTypeReader? reader = data == null ? null : OfficeOpenTypeReader.TryCreate(data);
            state = reader == null ? null : new RunKerningFont(reader);
            _fonts.Add(key, state);
            return state;
        }
    }

    private sealed class RunKerningFont {
        public RunKerningFont(OfficeOpenTypeReader reader) {
            Reader = reader;
            Kerning = OfficeOpenTypeKerning.FromReader(reader);
        }
        public OfficeOpenTypeReader Reader { get; }
        public OfficeOpenTypeKerning Kerning { get; }
        public int Scale(int value) => checked((int)Math.Round(value * 1000D / Reader.UnitsPerEm, MidpointRounding.AwayFromZero));
    }

    private static RichSeg CreateRunKerningProbe(string text, PdfTextRun run, PdfStandardFont baseFont, double fontSize, PdfCreateOptions? options) {
        PdfNamedFontFace? named = options != null && options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace face) ? face : null;
        return new RichSeg(text, run.Bold, run.Italic, false, false, null, null, null, null, null,
            ResolveFontForRun(run, baseFont), (run.FontSize?.ToPoints()) ?? fontSize, run.Baseline, 0D, namedFont: named,
            featureSettings: run.FeatureSettings, textDirection: run.TextDirection,
            horizontalTextScaling: run.HorizontalTextScaling, characterSpacing: (run.CharacterSpacing.ToPoints()),
            preservesPaintedGlyphs: run.PreservesPaintedGlyphs, formField: run.FormField);
    }
}
