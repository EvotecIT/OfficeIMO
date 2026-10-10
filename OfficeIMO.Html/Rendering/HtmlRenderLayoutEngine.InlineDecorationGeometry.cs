using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Records ordinary inline font boxes independently of the descendant paint
    /// extent. An atomic child's height and the containing line's extra leading
    /// must not become the non-replaced parent's decoration height.
    /// </summary>
    private void RecordInlineFontDecorationGeometry(InlineLine line, HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer, double lineStart, double lineY, double lineHeight, double baseline,
        bool emptyAtomicBaseline, bool bidiResolved,
        IReadOnlyDictionary<IElement, InlineContainingBounds> bounds) {
        bool qualifiedLine = !bidiResolved && line.HasFlowContent && !line.HasReplacedImage
            && paragraphStyle.WritingMode == "horizontal-tb" && paragraphStyle.Direction == "ltr"
            && line.Segments.All(segment => IsOrdinaryInlineDecorationRun(segment.Run, paragraphStyle));

        // Top-aligned atoms enlarge the line without lifting its font strut to
        // the atom's bottom. Baseline-aligned empty atoms already resolved the
        // shared baseline, including every participating parent strut.
        double decorationBaseline = baseline;
        bool hasAtomic = line.Segments.Any(segment => segment.Run.AtomicBlock != null);
        if (qualifiedLine && !emptyAtomicBaseline && hasAtomic) {
            decorationBaseline = ResolveEmptyAtomicStrutBaseline(paragraphStyle);
            foreach (HtmlRenderBoxStyle strut in line.Segments.SelectMany(segment => segment.Run.InlineStrutStyles)
                .Concat(line.EmptyInlineStrutStyles)) {
                decorationBaseline = Math.Max(decorationBaseline, ResolveEmptyAtomicStrutBaseline(strut));
            }
        }

        double cursor = lineStart;
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            double x = cursor + segment.LeadingAdvance;
            cursor += segment.Advance;
            if (run.IsFlowMarker || run.PositionedMarkerElement != null
                || run.RunningStringElement != null || run.RunningElementAssignment != null) continue;
            double segmentDecorationBaseline = decorationBaseline;
            if (qualifiedLine && !hasAtomic) {
                // Text and empty struts use the same selected-face placement,
                // including negative half-leading in lines shorter than an em.
                ResolveInlineTextVerticalPlacement(segment, false, lineY, lineHeight, baseline,
                    out double textY, out _, out _);
                segmentDecorationBaseline = textY - lineY + run.Style.Font.Size;
            }
            for (IElement? current = run.OwnerElement; current != null; current = current.ParentElement) {
                if (bounds.TryGetValue(current, out InlineContainingBounds? ownerBounds)
                    && _layoutStyles.TryGetValue(current, out HtmlRenderBoxStyle? style)
                    && style.Display == "inline" && HasInlineBoxPaint(style)) {
                    HtmlTextFaceMetrics? face = qualifiedLine && style.TableVerticalAlignment == "baseline"
                        && IsOrdinaryInlineDecorationStyle(style, paragraphStyle)
                        ? ResolveTextFaceMetrics(string.Empty, style) : null;
                    if (!IsUsableInlineDecorationFace(face)) {
                        // Specialized, mixed-face, bidi and displaced runs retain
                        // their existing route. Do not replace their geometry by
                        // assuming this ordinary horizontal baseline policy.
                        ownerBounds.RetainDecorationFallback();
                    } else {
                        double contentX = x;
                        double contentWidth = segment.Width;
                        HtmlInlineEdgeScope? scope = run.InlineEdgeScopes.FirstOrDefault(edge => ReferenceEquals(edge.Owner, current));
                        if (scope != null && _currentInlineEdgeGeometry.TryGetValue(scope, out var edges)) {
                            contentX = edges.Left;
                            contentWidth = Math.Max(0D, edges.Right - edges.Left);
                        }
                        ownerBounds.IncludeFontDecoration(contentX,
                            lineY + segmentDecorationBaseline - face!.Value.BaselineOffset,
                            contentWidth, face.Value.Height);
                    }
                }
                if (ReferenceEquals(current, formattingContainer)) break;
            }
        }
    }

    private static bool IsOrdinaryInlineDecorationRun(HtmlInlineRun run, HtmlRenderBoxStyle paragraphStyle) {
        if (run.IsFlowMarker || run.PositionedMarkerElement != null
            || run.RunningStringElement != null || run.RunningElementAssignment != null) return true;
        // Leaders paint with their own line-relative glyph/stroke placement.
        // Keep that specialized route outside the ordinary text/strut font box.
        if (run.LeaderPattern != null) return false;
        if (run.PaintOffsetX != 0D || run.PaintOffsetY != 0D
            || !IsOrdinaryInlineDecorationStyle(run.Style, paragraphStyle)
            || run.InlineStrutStyles.Any(style => !IsOrdinaryInlineDecorationStyle(style, paragraphStyle))) return false;
        return run.AtomicBlock == null ? run.Style.TableVerticalAlignment == "baseline"
            : run.HasEmptyInlineBlockBaseline && run.Style.Display == "inline-block"
                && run.Style.TableVerticalAlignment is "baseline" or "top";
    }

    private static bool IsOrdinaryInlineDecorationStyle(HtmlRenderBoxStyle style, HtmlRenderBoxStyle paragraphStyle) =>
        style.WritingMode == "horizontal-tb" && style.Direction == "ltr"
        && HasSameEmptyAtomicStrutFace(style, paragraphStyle) && style.Font.Size > 0D;

    private static bool IsUsableInlineDecorationFace(HtmlTextFaceMetrics? face) =>
        face.HasValue && face.Value.Height > 0D && !double.IsNaN(face.Value.Height) && !double.IsInfinity(face.Value.Height)
        && !double.IsNaN(face.Value.BaselineOffset) && !double.IsInfinity(face.Value.BaselineOffset)
        && face.Value.BaselineOffset >= 0D && face.Value.BaselineOffset <= face.Value.Height;
}
