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
            if (run.AtomicBlock != null) {
                if (!run.HasEmptyInlineBlockBaseline || run.Style.TableVerticalAlignment is not ("baseline" or "top")) return false;
            } else if (!HasSameEmptyAtomicStrutFace(run.Style, paragraphStyle)) return false;
            if (run.InlineEdgeScopes.Any(scope => !HasSameEmptyAtomicStrutFace(scope.Style, paragraphStyle))) return false;
        }

        double ascent = ResolveEmptyAtomicStrutBaseline(paragraphStyle);
        double descent = Math.Max(0D, paragraphStyle.LineHeight - ascent);
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (run.AtomicBlock != null) {
                if (run.Style.TableVerticalAlignment == "baseline") ascent = Math.Max(ascent, run.AtomicBlock.Height);
            } else if (!run.IsFlowMarker && run.RunningStringElement == null && run.RunningElementAssignment == null) {
                AddStrut(run.Style);
            }
            foreach (HtmlInlineEdgeScope scope in run.InlineEdgeScopes) AddStrut(scope.Style);
        }
        baseline = ascent;
        lineHeight = Math.Max(lineHeight, ascent + descent);
        return true;

        void AddStrut(HtmlRenderBoxStyle style) {
            double sourceBaseline = ResolveEmptyAtomicStrutBaseline(style);
            ascent = Math.Max(ascent, sourceBaseline);
            descent = Math.Max(descent, style.LineHeight - sourceBaseline);
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
