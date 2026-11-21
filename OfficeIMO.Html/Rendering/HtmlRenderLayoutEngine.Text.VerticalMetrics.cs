using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ResolveInlineTextVerticalPlacement(
        InlineSegment segment,
        bool alignBaseline,
        double lineY,
        double lineHeight,
        double baseline,
        bool useOrdinaryFaceMetrics,
        out double textY,
        out double paintHeight,
        out double paintTopOverflow) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        double fontSize = style.Font.Size;
        textY = alignBaseline
            // Positioned text paints its alphabetic baseline one source em below Y.
            // The line ascent determines the shared baseline, not the paint-frame origin.
            ? lineY + baseline - fontSize
            : lineY + Math.Min(0D, (lineHeight - fontSize) / 2D);
        paintHeight = Math.Max(lineHeight, fontSize);
        paintTopOverflow = 0D;
        HtmlTextFaceMetrics? face = useOrdinaryFaceMetrics && style.Font.Size > 0D
            ? ResolveTextFaceMetrics(segment.Text, style)
            : ResolveInlineShortLineFace(segment);
        if (!IsUsableTextFaceMetrics(face)) return;
        if (!face.HasValue) return;

        // CSS centers the face's ascent/descent box in the used line height. Positioned Drawing
        // text anchors its baseline one em below Y, so translate that browser baseline back to Y.
        if (!alignBaseline) {
            textY = useOrdinaryFaceMetrics && TryResolveNormalInlineTextMetrics(style, face, out InlineTextLineMetrics metrics)
                ? lineY + metrics.Baseline - fontSize
                : lineY + (lineHeight - face.Value.Height) / 2D
                    + face.Value.BaselineOffset - fontSize;
        }
        // This frame starts one source em above the baseline, rather than at the
        // face ascent. Its lower edge must include the selected face's descent.
        paintHeight = Math.Max(lineHeight, fontSize + face.Value.Height - face.Value.BaselineOffset);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - fontSize);
    }

    private double ResolveInlineSharedBaseline(InlineLine line, HtmlRenderBoxStyle paragraphStyle, ref double lineHeight) {
        bool hasImage = line.HasReplacedImage;
        bool useNormalImageMetrics = UsesNormalImageLineMetrics(line, paragraphStyle);
        string strutText = line.Segments.FirstOrDefault(segment => segment.Run.AtomicBlock == null
            && segment.Text.Length > 0)?.Text ?? string.Empty;
        InlineTextLineMetrics strut = ResolveInlineSharedTextMetrics(strutText, paragraphStyle, hasImage, useNormalImageMetrics);
        double baseline = strut.Baseline;
        double descent = Math.Max(0D, strut.Height - baseline);
        foreach (InlineSegment segment in line.Segments) {
            if (!InlineLine.ContributesFlowContent(segment.Run)) continue;
            HtmlRenderFlowBlock? atomic = segment.Run.AtomicBlock;
            if (atomic != null && !hasImage) continue;
            InlineTextLineMetrics source = atomic != null
                ? new InlineTextLineMetrics(atomic.Height,
                    Math.Min(atomic.Height, Math.Max(0D, segment.Run.AtomicBaseline ?? atomic.Height)))
                : ResolveInlineSharedTextMetrics(segment.Text, segment.Run.Style, hasImage, useNormalImageMetrics);
            double sourceBaseline = source.Baseline;
            baseline = Math.Max(baseline, sourceBaseline);
            descent = Math.Max(descent, source.Height - sourceBaseline);
        }
        // Replace the preliminary em-based estimate with the same used metrics
        // that align the image. Keeping its old descent as a minimum would leave
        // text-containing image lines taller than their image-only equivalents.
        if (hasImage) lineHeight = useNormalImageMetrics
            ? Math.Max(0.01D, baseline + descent)
            : Math.Max(lineHeight, baseline + descent);
        return baseline;
    }

    private double ResolveInlineTextBaseline(string text, HtmlRenderBoxStyle style, bool hasImage) {
        if (!hasImage && style.Font.Size > 0D
            && TryResolveNormalInlineTextMetrics(style, ResolveTextFaceMetrics(text, style), out InlineTextLineMetrics normal))
            return normal.Baseline;
        HtmlTextFaceMetrics? face = ResolveInlineShortLineFace(text, style);
        // Alignment and ink overflow are separate. The paragraph strut and each
        // text run use the same selected face, including beside an actual image.
        return face.HasValue
            ? (style.LineHeight - face.Value.Height) / 2D + face.Value.BaselineOffset
            : hasImage ? ResolveTextAscent(style) : style.Font.Size;
    }

    private HtmlTextFaceMetrics? ResolveInlineShortLineFace(InlineSegment segment) =>
        ResolveInlineShortLineFace(segment.Text, segment.Run.Style);

    private HtmlTextFaceMetrics? ResolveInlineShortLineFace(string text, HtmlRenderBoxStyle style) {
        if (style.LineHeight >= style.Font.Size) return null;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
        return IsUsableTextFaceMetrics(face) ? face : null;
    }

    private static bool IsUsableTextFaceMetrics(HtmlTextFaceMetrics? face) =>
        face.HasValue && face.Value.IsUsable;

    /// <summary>Resolves horizontal CSS leading from used font boxes while preserving the program's raw metrics and glyph ink.</summary>
    private static bool TryResolveNormalInlineTextMetrics(HtmlRenderBoxStyle style,
        HtmlTextFaceMetrics? face, out InlineTextLineMetrics metrics, bool allowShortLine = false) {
        metrics = default;
        if (style.Font.Size <= 0D || !allowShortLine && style.LineHeight < style.Font.Size
            || !TryResolveUsedInlineFontBox(face, out InlineTextLineMetrics fontBox)) return false;
        if (style.UsesNaturalLineHeight) {
            var natural = face!.Value.GetNaturalLineBox();
            metrics = new InlineTextLineMetrics(natural.Ascent + natural.Descent, natural.Ascent);
            return true;
        }
        double ascent = fontBox.Baseline;
        double descent = fontBox.Height - ascent;
        double usedHeight = style.InlineAlignmentLineHeight
            ?? HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(style.LineHeight, style.Font.Size);
        double leadingUnits = (usedHeight - ascent - descent) * 64D;
        // Divide signed fixed units before flooring the ascent share to a whole
        // pixel; the remaining leading belongs below the alphabetic baseline.
        double ascentLeading = Math.Floor(Math.Truncate(leadingUnits / 2D) / 64D);
        metrics = new InlineTextLineMetrics(usedHeight, ascent + ascentLeading);
        return true;
    }

    // Line metrics, paint placement and decoration geometry ask for the same word's face metrics
    // several times per line. Registered-face answers depend only on the text, the font request
    // and the registered faces, so they are kept until the face collection changes. Fallback
    // answers come from the host and keep its own reuse rules.
    private const int MaximumRememberedTextFaceMetrics = 16384;
    private Dictionary<TextFaceMetricsKey, HtmlTextFaceMetrics>? _textFaceMetrics;
    private object? _textFaceMetricsGeneration;

    private HtmlTextFaceMetrics? ResolveTextFaceMetrics(string text, HtmlRenderBoxStyle style) {
        object generation = _fonts.SelectionGeneration;
        if (!ReferenceEquals(generation, _textFaceMetricsGeneration)) {
            _textFaceMetrics?.Clear();
            _textFaceMetricsGeneration = generation;
            _selectableFacesFamily = null;
        }
        if (text.Length > 0 && IsUsableFontSize(style.Font.Size)
            && !HasSelectableFaces(style.Font.FamilyName, style.FontDescriptor)) {
            return _options.FallbackTextFaceMetrics?.Invoke(text, style.Font, style.FontDescriptor);
        }
        var key = new TextFaceMetricsKey(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size);
        if (_textFaceMetrics != null && _textFaceMetrics.TryGetValue(key, out HtmlTextFaceMetrics remembered)) return remembered;
        if (TryResolveRegisteredTextFaceMetrics(text, style, out HtmlTextFaceMetrics resolved)) {
            _textFaceMetrics ??= new Dictionary<TextFaceMetricsKey, HtmlTextFaceMetrics>();
            if (_textFaceMetrics.Count < MaximumRememberedTextFaceMetrics) _textFaceMetrics[key] = resolved;
            return resolved;
        }
        return _options.FallbackTextFaceMetrics?.Invoke(text, style.Font, style.FontDescriptor);
    }

    // A family list without any registered candidate face resolves every text to the host
    // fallback. Runs of one style share the family string, so the last answer is kept.
    private string? _selectableFacesFamily;
    private OfficeFontFaceDescriptor _selectableFacesDescriptor;
    private bool _selectableFaces;

    private bool HasSelectableFaces(string? familyNames, OfficeFontFaceDescriptor descriptor) {
        if (_selectableFacesFamily == null || !ReferenceEquals(_selectableFacesFamily, familyNames)
            || !_selectableFacesDescriptor.Equals(descriptor)) {
            _selectableFaces = _fonts.HasSelectableFaces(familyNames, descriptor);
            _selectableFacesFamily = familyNames;
            _selectableFacesDescriptor = descriptor;
        }
        return _selectableFaces;
    }

    private static bool IsUsableFontSize(double size) => size > 0D && !double.IsNaN(size) && !double.IsInfinity(size);

    private bool TryResolveRegisteredTextFaceMetrics(string text, HtmlRenderBoxStyle style, out HtmlTextFaceMetrics resolved) {
        resolved = default;
        if (text.Length == 0) {
            IOfficeFontProgram? strut = ResolveTextFaceProgram(text, style);
            if (strut is not IOfficeFontBaselineMetrics strutBaseline) return false;
            resolved = Metrics(strut, strutBaseline, style.Font.Size);
            return true;
        }
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, style.Font.FamilyName, style.FontDescriptor);
        HtmlTextFaceMetrics[]? metrics = null;
        for (int index = 0; index < runs.Count; index++) {
            OfficeFontFallbackRun run = runs[index];
            if (!_fonts.TryResolveFaceForText(run.Text, run.FamilyName, style.FontDescriptor, style.Font.Size, out OfficeFontFace? face)
                || face?.Program is not IOfficeFontBaselineMetrics baseline) {
                return false;
            }
            metrics ??= new HtmlTextFaceMetrics[runs.Count];
            metrics[index] = Metrics(face.Program, baseline, style.Font.Size);
        }
        HtmlTextFaceMetrics? combined = HtmlTextFaceMetrics.Combine(metrics ?? Array.Empty<HtmlTextFaceMetrics>());
        if (!combined.HasValue) return false;
        resolved = combined.Value;
        return true;

        static HtmlTextFaceMetrics Metrics(IOfficeFontProgram program, IOfficeFontBaselineMetrics baseline, double size) =>
            new(program.LineHeight(size), baseline.BaselineOffset(size), program.LineSpacingRatio * size);
    }

    private readonly struct TextFaceMetricsKey : IEquatable<TextFaceMetricsKey> {
        private readonly string _text;
        private readonly string? _family;
        private readonly OfficeFontFaceDescriptor _descriptor;
        private readonly long _sizeBits;

        internal TextFaceMetricsKey(string text, string? family, OfficeFontFaceDescriptor descriptor, double size) {
            _text = text;
            _family = family;
            _descriptor = descriptor;
            _sizeBits = BitConverter.DoubleToInt64Bits(size);
        }

        public bool Equals(TextFaceMetricsKey other) =>
            _sizeBits == other._sizeBits
            && _descriptor.Equals(other._descriptor)
            && string.Equals(_text, other._text, StringComparison.Ordinal)
            && string.Equals(_family, other._family, StringComparison.Ordinal);

        public override bool Equals(object? obj) => obj is TextFaceMetricsKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.Ordinal.GetHashCode(_text);
                hash = (hash * 397) ^ (_family == null ? 0 : StringComparer.Ordinal.GetHashCode(_family));
                hash = (hash * 397) ^ _descriptor.GetHashCode();
                return (hash * 397) ^ _sizeBits.GetHashCode();
            }
        }
    }

    private IOfficeFontProgram? ResolveTextFaceProgram(string text, HtmlRenderBoxStyle style) =>
        text.Length == 0
            ? _fonts.ResolveForText(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size, out _)
            : _fonts.TryResolveFaceForText(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size, out OfficeFontFace? face)
                ? face?.Program : null;

    private void ResolveInlineAnchorTextVerticalBounds(
        InlineSegment segment,
        double lineY,
        double lineHeight,
        double textY,
        out double anchorY,
        out double anchorHeight) {
        anchorY = lineY;
        anchorHeight = lineHeight;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(segment.Text, segment.Run.Style);
        if (!face.HasValue || face.Value.Height <= 0D || double.IsNaN(face.Value.Height)
            || double.IsInfinity(face.Value.Height)) return;

        // An inline anchor's CSS box follows its font's ascent/descent box, not
        // the full line box. Extra leading belongs to the line around it.
        anchorHeight = face.Value.Height;
        if (TryResolveNormalInlineTextMetrics(segment.Run.Style, face, out _)) {
            anchorY = textY + segment.Run.Style.Font.Size - Math.Floor(face.Value.BaselineOffset + 0.5D);
        }
    }

    private readonly struct InlineTextLineMetrics {
        internal InlineTextLineMetrics(double height, double baseline) {
            Height = height;
            Baseline = baseline;
        }

        internal double Height { get; }
        internal double Baseline { get; }
    }
}
