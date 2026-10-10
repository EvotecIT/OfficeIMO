namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Resolves the synthesized bottom-edge baseline of ordinary empty inline-blocks
    /// against the containing line's strut. Content-derived, mixed-face and vertical
    /// baselines retain their existing route and qualification boundary.
    /// </summary>
    private bool TryResolveEmptyInlineBlockLineMetrics(InlineLine line, HtmlRenderBoxStyle paragraphStyle,
        ref double lineHeight, out double baseline) {
        baseline = 0D;
        if (!line.HasFlowContent || IsVerticalWritingMode(paragraphStyle.WritingMode)
            || !line.Segments.Any(segment => segment.Run.HasEmptyInlineBlockBaseline
                && segment.Run.Style.TableVerticalAlignment == "baseline")) return false;

        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (run.IsFlowMarker) continue;
            if (run.AtomicBlock != null) {
                if (!run.HasEmptyInlineBlockBaseline || run.Style.TableVerticalAlignment is not ("baseline" or "top")) return false;
            } else if (!HasSameEmptyAtomicStrutFace(run.Style, paragraphStyle)) return false;
            if (run.InlineStrutStyles.Any(style => !HasSameEmptyAtomicStrutFace(style, paragraphStyle))) return false;
        }
        if (line.EmptyInlineStrutStyles.Any(style => !HasSameEmptyAtomicStrutFace(style, paragraphStyle))) return false;

        double ascent = ResolveEmptyAtomicStrutBaseline(paragraphStyle);
        double descent = Math.Max(0D, paragraphStyle.LineHeight - ascent);
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (run.IsFlowMarker) continue;
            if (run.AtomicBlock != null) {
                if (run.Style.TableVerticalAlignment == "baseline") ascent = Math.Max(ascent, run.AtomicBlock.Height);
            } else if (run.RunningStringElement == null && run.RunningElementAssignment == null) {
                AddStrut(run.Style);
            }
            foreach (HtmlRenderBoxStyle style in run.InlineStrutStyles) AddStrut(style);
        }
        foreach (HtmlRenderBoxStyle style in line.EmptyInlineStrutStyles) AddStrut(style);
        baseline = ascent;
        lineHeight = Math.Max(lineHeight, ascent + descent);
        return true;

        void AddStrut(HtmlRenderBoxStyle style) {
            double sourceBaseline = ResolveEmptyAtomicStrutBaseline(style);
            ascent = Math.Max(ascent, sourceBaseline);
            descent = Math.Max(descent, style.LineHeight - sourceBaseline);
        }
    }

    private sealed partial class InlineLine {
        private List<HtmlRenderBoxStyle>? _emptyInlineStrutStyles;
        internal IReadOnlyList<HtmlRenderBoxStyle> EmptyInlineStrutStyles =>
            _emptyInlineStrutStyles ?? (IReadOnlyList<HtmlRenderBoxStyle>)Array.Empty<HtmlRenderBoxStyle>();

        // These boxes participate in an existing atomic line without creating
        // in-flow content, horizontal advances or a line on their own.
        internal void RecordEmptyInlineStruts(IReadOnlyList<HtmlRenderBoxStyle> styles) {
            if (styles.Count == 0) return;
            (_emptyInlineStrutStyles ??= new List<HtmlRenderBoxStyle>()).AddRange(styles);
        }

        // A nowrap suffix carries its zero-width inline boxes to the same line
        // as its content; remove them before calculating the preceding line.
        internal IReadOnlyList<HtmlRenderBoxStyle> TakeEmptyInlineStruts(int start) {
            if (_emptyInlineStrutStyles == null || start >= _emptyInlineStrutStyles.Count) {
                return Array.Empty<HtmlRenderBoxStyle>();
            }
            HtmlRenderBoxStyle[] styles = _emptyInlineStrutStyles.Skip(start).ToArray();
            _emptyInlineStrutStyles.RemoveRange(start, _emptyInlineStrutStyles.Count - start);
            return styles;
        }
    }

    private static void CaptureInlineStrutStyles(IList<HtmlInlineRun> runs, int firstInlineRun, HtmlRenderBoxStyle style) {
        for (int index = firstInlineRun; index < runs.Count; index++) {
            HtmlInlineRun run = runs[index];
            run.InlineStrutStyles = run.InlineStrutStyles.Count == 0
                ? new[] { style }
                : run.InlineStrutStyles.Concat(new[] { style }).ToArray();
        }
    }

    private static bool HasSameEmptyAtomicStrutFace(HtmlRenderBoxStyle style, HtmlRenderBoxStyle paragraphStyle) =>
        string.Equals(style.Font.FamilyName, paragraphStyle.Font.FamilyName, StringComparison.OrdinalIgnoreCase)
        && Math.Abs(style.Font.Size - paragraphStyle.Font.Size) < 0.000001D
        && style.FontDescriptor.Equals(paragraphStyle.FontDescriptor)
        && style.BaselineScale == 1D && style.BaselineOffset == 0D;

    private double ResolveEmptyAtomicStrutBaseline(HtmlRenderBoxStyle style) {
        IOfficeFontProgram? program = _fonts.ResolveForText(string.Empty, style.Font.FamilyName,
            style.FontDescriptor, style.Font.Size, out _);
        HtmlTextFaceMetrics? face = program is IOfficeFontBaselineMetrics metrics
            ? new HtmlTextFaceMetrics(program.LineHeight(style.Font.Size), metrics.BaselineOffset(style.Font.Size))
            : _options.FallbackTextFaceMetrics?.Invoke(string.Empty, style.Font, style.FontDescriptor);
        if (face.HasValue && face.Value.Height > 0D && !double.IsNaN(face.Value.Height) && !double.IsInfinity(face.Value.Height)
            && !double.IsNaN(face.Value.BaselineOffset) && !double.IsInfinity(face.Value.BaselineOffset)
            && face.Value.BaselineOffset >= 0D && face.Value.BaselineOffset <= face.Value.Height) {
            return (style.LineHeight - face.Value.Height) / 2D + face.Value.BaselineOffset;
        }
        return ResolveTextAscent(style);
    }
}
