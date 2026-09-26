using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    private static List<Label> GetLabels(PdfLogicalPage page, IReadOnlyList<PdfStaticFormTextEvidence> ocrText,
        double pageWidth, double pageHeight,
        IReadOnlyList<PaintArea> filledAreas, IReadOnlyList<PdfPageDrawingEffectTransition> effects, IReadOnlyList<VisualRect> tableBounds, IReadOnlyList<PdfPageVisualPrimitive> primitives,
        out List<VisualRect> nativeTextBounds, ref long candidateScanWork, int maxCandidateScanWork,
        CancellationToken cancellationToken) {
        var labels = new List<Label>();
        nativeTextBounds = new List<VisualRect>();
        foreach (PdfLogicalTextBlock block in page.TextBlocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block.XEnd <= block.XStart) continue;
            PdfLogicalVisualBounds bounds = block.VisualBounds ?? ToVisualBounds(page, block);
            bool completeBlock = Valid(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, pageWidth, pageHeight);
            if (!TryIntersectPage(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, pageWidth, pageHeight, out VisualRect visual)) continue;
            bool visibleText = block.Spans.Count == 0;
            bool provableLabel = block.Spans.Count == 0 && completeBlock;
            string text = NormalizeLabel(block.Text);
            if (block.Spans.Count > 0) {
                bool uncertainEffect = false;
                var visibleSpans = new List<PdfTextSpan>(block.Spans.Count);
                VisualRect? labelBounds = null;
                foreach (PdfTextSpan span in block.Spans) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!span.IsVisible || (span.Color?.A ?? 255) <= 3 && !span.HasAdditionalVisiblePaint) continue;
                    PdfTextSpanBounds spanBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
                    if (span.VisibleStrokePadding > 0D) {
                        double padding = span.VisibleStrokePadding;
                        spanBounds = new PdfTextSpanBounds(spanBounds.Left - padding, spanBounds.Bottom - padding,
                            spanBounds.Right + padding, spanBounds.Top + padding);
                    }
                    bool fullyVisibleText = true;
                    if (span.ClipPath is PdfPageClipPath clip) {
                        var paintedText = PdfPageClipPath.Rectangle(spanBounds.Left,
                            page.Height - spanBounds.Top, spanBounds.Right - spanBounds.Left,
                            spanBounds.Top - spanBounds.Bottom);
                        if (clip.Width <= 0D || clip.Height <= 0D ||
                            clip.CanProveNoPositiveAreaIntersection(paintedText)) continue;
                        fullyVisibleText = clip.IsRectangle && clip.IsExact && !clip.ContainsTextClipping &&
                            clip.X <= paintedText.X && clip.Y <= paintedText.Y &&
                            clip.X + clip.Width >= paintedText.X + paintedText.Width &&
                            clip.Y + clip.Height >= paintedText.Y + paintedText.Height;
                    }
                    PdfVisualBounds spanProjected = page.TransformBoundsToVisual(spanBounds.Left, spanBounds.Bottom,
                        spanBounds.Right, spanBounds.Top);
                    fullyVisibleText &= Valid(spanProjected.Left, spanProjected.Top, spanProjected.Right, spanProjected.Bottom,
                        pageWidth, pageHeight);
                    if (!TryIntersectPage(spanProjected.Left, spanProjected.Top, spanProjected.Right, spanProjected.Bottom,
                        pageWidth, pageHeight, out VisualRect spanVisual)) continue;
                    if (span.ClipPath is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } exactClip) {
                        PdfVisualBounds clipped = page.TransformBoundsToVisual(exactClip.X,
                            page.Height - exactClip.Y - exactClip.Height, exactClip.X + exactClip.Width, page.Height - exactClip.Y);
                        spanVisual = new VisualRect(Math.Max(spanVisual.Left, clipped.Left), Math.Max(spanVisual.Top, clipped.Top),
                            Math.Min(spanVisual.Right, clipped.Right), Math.Min(spanVisual.Bottom, clipped.Bottom));
                        if (spanVisual.Right <= spanVisual.Left || spanVisual.Bottom <= spanVisual.Top) continue;
                    }
                    bool uncertainSpan = !fullyVisibleText || span.HasUnresolvedPaint || span.HasAdditionalVisiblePaint ||
                        (span.Color?.A ?? 255) < 46;
                    if (uncertainSpan) uncertainEffect = true;
                    ChargeEffectLookup(effects.Count, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
                    PdfPageDrawingEffect effect = PdfReadPage.ResolveDrawingEffect(effects, span.PaintOrder,
                        contentOrderKey: span.ContentOrderKey);
                    if (effect.HasUnsupportedGraphicsState || effect.SoftMask is not null || effect.HasUnresolvedSoftMask ||
                        effect.BlendMode != OfficeBlendMode.Normal) uncertainEffect = true;
                    bool covered = false;
                    foreach (PaintArea area in filledAreas) {
                        cancellationToken.ThrowIfCancellationRequested();
                        candidateScanWork++;
                        if (candidateScanWork > maxCandidateScanWork) {
                            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                                maxCandidateScanWork, candidateScanWork);
                        }
                        if (IsLater(area.PaintOrder, area.ContentOrderKey, span.PaintOrder, span.ContentOrderKey) &&
                            OverlapArea(area.Bounds, spanVisual) > 0D) uncertainEffect = true;
                        if (IsLaterCover(area, spanVisual, span.PaintOrder, span.ContentOrderKey)) {
                            covered = true;
                            break;
                        }
                    }
                    if (!covered) {
                        foreach (PdfLogicalImage image in page.Images) {
                            foreach (PdfImagePlacement placement in image.Placements) {
                                cancellationToken.ThrowIfCancellationRequested();
                                candidateScanWork++;
                                if (candidateScanWork > maxCandidateScanWork) {
                                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                                        maxCandidateScanWork, candidateScanWork);
                                }
                                if (IsLaterOpaqueImageCover(page, image, placement, spanVisual,
                                    span.PaintOrder, span.ContentOrderKey)) { covered = true; uncertainEffect = true; break; }
                                if (HasUnprovenImageOverlap(page, placement, spanVisual)) uncertainEffect = true;
                            }
                            if (covered) break;
                        }
                    }
                    if (!covered) {
                        ChargeEffectLookup(primitives.Count, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
                        foreach (PdfPageVisualPrimitive paint in primitives) {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!IsLater(paint.PaintOrder, paint.ContentOrderKey, span.PaintOrder, span.ContentOrderKey) ||
                                !TryGetStrokeBounds(paint, out VisualRect stroke, out VisualRect? hollow)) continue;
                            if (OverlapArea(stroke, spanVisual) - (hollow is VisualRect hole ? OverlapArea(hole, spanVisual) : 0D) > 0D) {
                                uncertainEffect = true;
                                break;
                            }
                        }
                    }
                    if (covered) continue;
                    if (uncertainSpan) {
                        nativeTextBounds.Add(spanVisual);
                        continue;
                    }
                    if (span.IsArtifactContent) {
                        nativeTextBounds.Add(spanVisual);
                        continue;
                    }
                    if (span.Color is OfficeColor ink &&
                        !HasContrastingBackdrop(filledAreas, spanVisual, ink, span.PaintOrder, span.ContentOrderKey)) {
                        // Keep the text as possible field occupancy, but do not infer a label without visible contrast.
                        nativeTextBounds.Add(spanVisual);
                        uncertainEffect = true;
                        continue;
                    }
                    visibleSpans.Add(span);
                    nativeTextBounds.Add(spanVisual);
                    labelBounds = labelBounds.HasValue
                        ? new VisualRect(Math.Min(labelBounds.Value.Left, spanVisual.Left),
                            Math.Min(labelBounds.Value.Top, spanVisual.Top),
                            Math.Max(labelBounds.Value.Right, spanVisual.Right),
                            Math.Max(labelBounds.Value.Bottom, spanVisual.Bottom))
                        : spanVisual;
                }
                visibleText = visibleSpans.Count > 0;
                provableLabel = visibleText && !uncertainEffect;
                if (labelBounds.HasValue) visual = labelBounds.Value;
                if (visibleSpans.Count != block.Spans.Count) {
                    text = NormalizeLabel(string.Concat(visibleSpans.Select(static span => span.Text)));
                }
            }
            if (visibleText && block.Spans.Count == 0) nativeTextBounds.Add(visual);
            if (!provableLabel || block.IsTableContent || text.Length == 0 || text.Length > 80) continue;
            ChargeEffectLookup(tableBounds.Count, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
            if (OverlapsDetectedTable(tableBounds, visual)) continue;
            labels.Add(new Label(text, visual, block.Confidence, isOcr: false));
        }
        foreach (PdfStaticFormTextEvidence item in ocrText) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.PageNumber != page.PageNumber) continue;
            string text = NormalizeLabel(item.Text);
            if (!Valid(item.Left, item.Top, item.Right, item.Bottom, pageWidth, pageHeight)) continue;
            var bounds = new VisualRect(item.Left, item.Top, item.Right, item.Bottom);
            nativeTextBounds.Add(bounds);
            if (text.Length == 0 || text.Length > 80) continue;
            ChargeEffectLookup(tableBounds.Count, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
            if (OverlapsDetectedTable(tableBounds, bounds)) continue;
            // Overlapping detections of the same text share one physical owner. Merge
            // their bounds before assignment, including evidence connected through a third detection.
            double confidence = item.Confidence;
            bool isOcr = true;
            bool merged;
            do {
                merged = false;
                for (int index = labels.Count - 1; index >= 0; index--) {
                    ChargeEffectLookup(0, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
                    Label existing = labels[index];
                    if (!string.Equals(existing.Text, text, StringComparison.OrdinalIgnoreCase) ||
                        OverlapArea(existing.Bounds, bounds) <= 0D) continue;
                    bounds = new VisualRect(Math.Min(bounds.Left, existing.Bounds.Left), Math.Min(bounds.Top, existing.Bounds.Top),
                        Math.Max(bounds.Right, existing.Bounds.Right), Math.Max(bounds.Bottom, existing.Bounds.Bottom));
                    double existingScore = GetLabelConfidence(existing.Confidence, existing.IsOcr);
                    double currentScore = GetLabelConfidence(confidence, isOcr);
                    if (existingScore > currentScore || existingScore == currentScore && !existing.IsOcr) {
                        confidence = existing.Confidence;
                        isOcr = existing.IsOcr;
                    }
                    labels.RemoveAt(index);
                    merged = true;
                }
            } while (merged && labels.Count > 0);
            ChargeEffectLookup(labels.Count, ref candidateScanWork, maxCandidateScanWork, cancellationToken);
            if (item.Confidence >= 0.8D) {
                // A high-confidence caller correction supersedes conflicting native extraction at the same location.
                labels.RemoveAll(label => !label.IsOcr &&
                    OverlapArea(label.Bounds, bounds) > Math.Min(label.Bounds.Area, bounds.Area) * 0.5D);
            }
            labels.Add(new Label(text, bounds, confidence, isOcr));
        }
        return labels;
    }


    // Partial native text still occupies the visible page. Full bounds are required only for labels.
    private static bool TryIntersectPage(double left, double top, double right, double bottom,
        double width, double height, out VisualRect bounds) {
        bounds = default;
        if (double.IsNaN(left) || double.IsNaN(top) || double.IsNaN(right) || double.IsNaN(bottom) ||
            double.IsInfinity(left) || double.IsInfinity(top) || double.IsInfinity(right) || double.IsInfinity(bottom) ||
            right <= left || bottom <= top) return false;
        bounds = new VisualRect(Math.Max(0D, left), Math.Max(0D, top), Math.Min(width, right), Math.Min(height, bottom));
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    // The canonical effect resolver may scan the entire timeline. Charge its
    // conservative upper bound before lookup, including pages with no candidates.
    private static void ChargeEffectLookup(int transitionCount, ref long work, int maximumWork,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        work = checked(work + transitionCount + 1L);
        if (work > maximumWork) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximumWork, work);
        }
    }

    // Image geometry does not prove its rendered color. Native labels over an image
    // need separate visibility evidence, such as caller-supplied positioned OCR.
    private static bool HasUnprovenImageOverlap(PdfLogicalPage page,
        PdfImagePlacement placement, VisualRect label) {
        if (placement.IsHiddenOptionalContent || placement.Opacity <= 0D ||
            placement.Width <= 0D || placement.Height <= 0D) return false;
        PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
            placement.X + placement.Width, placement.Y + placement.Height);
        var visible = new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
        if (placement.Clip is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } clip) {
            PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
            visible = new VisualRect(Math.Max(visible.Left, clipped.Left), Math.Max(visible.Top, clipped.Top),
                Math.Min(visible.Right, clipped.Right), Math.Min(visible.Bottom, clipped.Bottom));
        }
        return OverlapArea(visible, label) > 0D;
    }
}
