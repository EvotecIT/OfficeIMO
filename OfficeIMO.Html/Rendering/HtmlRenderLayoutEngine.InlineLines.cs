using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private sealed partial class InlineLine {
        private int _flowContentCount;
        private double _availableWidth;
        private double _indent;
        internal double IndentOffset { get; private set; }

        internal List<InlineSegment> Segments { get; } = new List<InlineSegment>();
        internal double Width { get; private set; }
        internal bool HasFlowContent => _flowContentCount > 0;
        internal bool HasForcedBreak { get; set; }
        internal bool HasExplicitPlacement { get; private set; }
        internal double X { get; private set; }
        internal double Y { get; private set; }
        internal double AvailableWidth => Math.Max(0.01D, _availableWidth - _indent);
        internal HtmlRenderBoxStyle? ParagraphStyle { get; set; }
        internal bool EndsWithHyphenation { get; set; }

        internal void Place(double x, double y, double availableWidth) {
            HasExplicitPlacement = true;
            X = Math.Max(0D, x);
            Y = Math.Max(0D, y);
            _availableWidth = Math.Max(0.01D, availableWidth);
        }

        internal void Indent(double indent, bool rightToLeft) {
            _indent = indent;
            IndentOffset = rightToLeft ? 0D : indent;
        }

        internal double ResolveAvailableWidth(double width) => HasExplicitPlacement
            ? AvailableWidth : Math.Max(0.01D, width - _indent);

        internal double ResolveAlignmentOffset(OfficeTextAlignment alignment, double width) {
            if (_indent == 0D) return ResolveLineOffset(alignment, ResolveAvailableWidth(width), Width);
            // An indent can exceed the line width. Preserve the signed space for
            // alignment so RTL text overflows towards inline-end, not the right.
            double remaining = (HasExplicitPlacement ? _availableWidth : width) - _indent - Width;
            if (alignment == OfficeTextAlignment.Right) return remaining;
            if (alignment == OfficeTextAlignment.Center) return remaining / 2D;
            return 0D;
        }

        internal void Add(InlineSegment segment) {
            segment.LeadingAdvance = 0D;
            segment.TrailingAdvance = 0D;
            Segments.Add(segment);
            Width += segment.Width;
            AddInlineEdges(segment);
            if (ContributesFlowContent(segment.Run)) _flowContentCount++;
        }

        internal void RemoveAt(int index) {
            InlineSegment removed = Segments[index];
            if (ContributesFlowContent(removed.Run)) _flowContentCount--;
            Width -= Segments[index].Width;
            Segments.RemoveAt(index);
            if (_inlineEdges.Count > 0) RebuildInlineEdges(removed);
        }

        /// <summary>
        /// Removes fragments being moved to another line while preserving their
        /// scope content anchors, rebuilding the remaining line's edges once.
        /// </summary>
        internal void RemoveRange(int index, int count) {
            for (int offset = 0; offset < count; offset++) {
                InlineSegment removed = Segments[index + offset];
                if (ContributesFlowContent(removed.Run)) _flowContentCount--;
                Width -= removed.Width;
            }
            Segments.RemoveRange(index, count);
            if (_inlineEdges.Count > 0) RebuildInlineEdges();
        }

        // Out-of-flow markers keep their static position segment without
        // creating an in-flow line box or contributing a font strut.
        private static bool ContributesFlowContent(HtmlInlineRun run) =>
            run.RunningStringElement == null && run.RunningElementAssignment == null
            && run.PositionedMarkerElement == null && !run.IsFlowMarker;

        internal void SetSegmentWidth(int index, double width) {
            double normalized = Math.Max(0D, width);
            Width += normalized - Segments[index].Width;
            Segments[index].SetWidth(normalized);
        }

        internal double ResolveLineHeight(double fallback) {
            if (!HasFlowContent) return 0D;
            double height = fallback;
            for (int i = 0; i < Segments.Count; i++) {
                if (Segments[i].Run.PositionedMarkerElement != null) continue;
                height = Math.Max(height, Segments[i].Run.AtomicBlock?.Height ?? Segments[i].Run.Style.LineHeight);
            }
            if (!HasReplacedImage) {
                // Content-derived atomic baselines retain their separate route.
                // Text-only face metrics are resolved by the layout engine below.
                if (Segments.Any(segment => segment.Run.AtomicBlock != null && !segment.Run.IsFlowMarker)
                    && HasMixedTextSizes(Segments[0].Run.Style)) {
                    double textBaseline = 0D;
                    double textDescent = 0D;
                    foreach (InlineSegment segment in Segments) {
                        if (segment.Run.AtomicBlock == null && segment.Run.PositionedMarkerElement == null) {
                            textBaseline = Math.Max(textBaseline, segment.Run.Style.Font.Size);
                            textDescent = Math.Max(textDescent, segment.Run.Style.LineHeight - segment.Run.Style.Font.Size);
                        }
                    }
                    height = Math.Max(height, textBaseline + textDescent);
                }
                return Math.Max(0.01D, height);
            }

            double ascent = 0D;
            double descent = 0D;
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                if (run.PositionedMarkerElement != null) continue;
                if (run.AtomicBlock != null) {
                    double atomicBaseline = Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height));
                    ascent = Math.Max(ascent, atomicBaseline);
                    descent = Math.Max(descent, run.AtomicBlock.Height - atomicBaseline);
                } else {
                    ascent = Math.Max(ascent, ResolveTextAscent(run.Style));
                    descent = Math.Max(descent, Math.Max(0D, run.Style.LineHeight - ResolveTextAscent(run.Style)));
                }
            }
            return Math.Max(0.01D, Math.Max(height, ascent + descent));
        }

        internal bool HasReplacedImage => Segments.Any(segment => segment.Run.IsReplacedImage);

        internal bool HasMixedTextSizes(HtmlRenderBoxStyle paragraphStyle) =>
            Segments.Any(segment => segment.Run.AtomicBlock == null && segment.Run.PositionedMarkerElement == null
                && Math.Abs(segment.Run.Style.Font.Size - paragraphStyle.Font.Size) > 0.000001D);

        internal double ResolveBaseline(HtmlRenderBoxStyle paragraphStyle) {
            if (!HasReplacedImage) {
                // Positioned text in the shared drawing model paints its baseline
                // at Y + the source font size. Keep mixed-size text on that same
                // baseline, including the containing paragraph's line strut.
                double baseline = paragraphStyle.Font.Size;
                foreach (InlineSegment segment in Segments) {
                    if (segment.Run.AtomicBlock == null && segment.Run.PositionedMarkerElement == null)
                        baseline = Math.Max(baseline, segment.Run.Style.Font.Size);
                }
                return baseline;
            }
            double ascent = ResolveTextAscent(paragraphStyle);
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                if (run.PositionedMarkerElement != null) continue;
                ascent = Math.Max(ascent, run.AtomicBlock == null
                    ? ResolveTextAscent(run.Style)
                    : Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height)));
            }
            return ascent;
        }
    }

    private static double ResolveTextAscent(HtmlRenderBoxStyle style) {
        double effectiveSize = GetEffectiveTextFont(style).Size;
        double leading = Math.Max(0D, style.LineHeight - effectiveSize);
        return Math.Min(style.LineHeight, leading / 2D + effectiveSize * 0.8D);
    }

    private static OfficeFontInfo GetEffectiveTextFont(HtmlRenderBoxStyle style) =>
        Math.Abs(style.BaselineScale - 1D) < 0.000001D
            ? style.Font
            : style.Font.WithSize(style.Font.Size * style.BaselineScale);
}
