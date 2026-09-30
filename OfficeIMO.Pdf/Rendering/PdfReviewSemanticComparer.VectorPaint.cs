using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfReviewSemanticComparer {
    private static bool HasChangedVectorPaintInsideClassifiedBounds(
        PdfReadPage expectedReadPage, PdfReadPage actualReadPage,
        PdfLogicalPage expected, PdfLogicalPage actual,
        List<PdfPixelRegion> classified, PdfVisualPageComparison visual,
        PdfVisualComparisonOptions options, CancellationToken cancellationToken) {
        if (classified.Count == 0) return false;
        IReadOnlyList<PdfPageVisualPrimitive> before = expectedReadPage.GetIdentityVisualPrimitives(cancellationToken);
        IReadOnlyList<PdfPageVisualPrimitive> after = actualReadPage.GetIdentityVisualPrimitives(cancellationToken);
        long work = 0L;
        var matchedAfter = Enumerable.Repeat(-1, before.Count).ToArray();
        var usedAfter = new bool[after.Count];
        int pairedCount = Math.Min(before.Count, after.Count);
        for (int index = 0; index < pairedCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeVectorWork(ref work, 1L + before[index].PathCommands.Count + after[index].PathCommands.Count);
            if (!SameSimpleVectorPaint(before[index], after[index])) continue;
            matchedAfter[index] = index;
            usedAfter[index] = true;
        }
        var unmatchedAfter = new Dictionary<int, LinkedList<int>>();
        for (int index = 0; index < after.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (usedAfter[index] || !IsSimpleVectorPaint(after[index])) continue;
            int key = VectorBucket(after[index]);
            if (!unmatchedAfter.TryGetValue(key, out LinkedList<int>? bucket)) {
                bucket = new LinkedList<int>();
                unmatchedAfter.Add(key, bucket);
            }
            bucket.AddLast(index);
        }
        for (int index = 0; index < before.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (matchedAfter[index] >= 0 || !IsSimpleVectorPaint(before[index]) ||
                !unmatchedAfter.TryGetValue(VectorBucket(before[index]), out LinkedList<int>? bucket)) continue;
            for (LinkedListNode<int>? node = bucket.First; node != null; node = node.Next) {
                ChargeVectorWork(ref work, 1L + before[index].PathCommands.Count + after[node.Value].PathCommands.Count);
                if (!SameSimpleVectorPaint(before[index], after[node.Value])) continue;
                matchedAfter[index] = node.Value;
                usedAfter[node.Value] = true;
                bucket.Remove(node);
                break;
            }
        }
        for (int index = 0; index < before.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (matchedAfter[index] >= 0) {
                if (RelativePaintOrderDiffers(before[index], after[matchedAfter[index]], expected, actual,
                        classified, visual, options, ref work, cancellationToken) &&
                    (ChangedPixelsIntersectVector(before[index], expected, classified, visual, options, ref work, cancellationToken) ||
                     ChangedPixelsIntersectVector(after[matchedAfter[index]], actual, classified, visual, options, ref work, cancellationToken))) return true;
            } else if (ChangedPixelsIntersectVector(before[index], expected,
                           classified, visual, options, ref work, cancellationToken)) return true;
        }
        for (int index = 0; index < after.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!usedAfter[index] && ChangedPixelsIntersectVector(after[index], actual,
                    classified, visual, options, ref work, cancellationToken)) return true;
        }
        int highestAfter = -1;
        bool reordered = false;
        for (int index = 0; index < matchedAfter.Length; index++) {
            if (matchedAfter[index] < 0) continue;
            if (matchedAfter[index] < highestAfter) { reordered = true; break; }
            highestAfter = matchedAfter[index];
        }
        if (reordered) {
            for (int first = 0; first < before.Count; first++) {
                if (matchedAfter[first] < 0) continue;
                cancellationToken.ThrowIfCancellationRequested();
                for (int second = first + 1; second < before.Count; second++) {
                    ChargeVectorWork(ref work, 1L);
                    if (matchedAfter[second] < 0 || matchedAfter[first] < matchedAfter[second]) continue;
                    if (!VectorBoundsOverlap(before[first], before[second])) continue;
                    if (ChangedPixelsIntersectVector(before[first], expected, classified, visual, options, ref work, cancellationToken) ||
                        ChangedPixelsIntersectVector(before[second], expected, classified, visual, options, ref work, cancellationToken)) return true;
                }
            }
        }
        return false;
    }

    private static int VectorBucket(PdfPageVisualPrimitive primitive) {
        unchecked {
            int hash = (int)primitive.Kind;
            hash = hash * 31 + primitive.X.GetHashCode();
            hash = hash * 31 + primitive.Y.GetHashCode();
            hash = hash * 31 + primitive.Width.GetHashCode();
            return hash * 31 + primitive.Height.GetHashCode();
        }
    }

    private static void ChargeVectorWork(ref long work, long amount) {
        const long maxWork = 200_000_000L;
        work = checked(work + amount);
        if (work > maxWork) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maxWork, work);
    }

    private static bool VectorBoundsOverlap(PdfPageVisualPrimitive first, PdfPageVisualPrimitive second) {
        double firstStroke = Math.Max(0D, first.StrokeWidth / 2D);
        double secondStroke = Math.Max(0D, second.StrokeWidth / 2D);
        return first.X - firstStroke < second.X + second.Width + secondStroke &&
            second.X - secondStroke < first.X + first.Width + firstStroke &&
            first.Y - firstStroke < second.Y + second.Height + secondStroke &&
            second.Y - secondStroke < first.Y + first.Height + firstStroke;
    }

    private static bool RelativePaintOrderDiffers(PdfPageVisualPrimitive before, PdfPageVisualPrimitive after,
        PdfLogicalPage expected, PdfLogicalPage actual, List<PdfPixelRegion> classified,
        PdfVisualPageComparison visual, PdfVisualComparisonOptions options,
        ref long work, CancellationToken cancellationToken) {
        double stroke = Math.Max(0D, before.StrokeWidth / 2D);
        var vectorBounds = new PdfLogicalVisualBounds(before.X - stroke, before.Y - stroke,
            before.X + before.Width + stroke, before.Y + before.Height + stroke);
        if (!ToPixelRegion(vectorBounds, expected, visual, options, out PdfPixelRegion vectorRegion) ||
            visual.ChangedBounds is not PdfPixelRegion changedBounds ||
            !PixelRegionsOverlap(vectorRegion, changedBounds)) return false;
        bool intersectsClassified = false;
        foreach (PdfPixelRegion region in classified) {
            ChargeVectorWork(ref work, 1L);
            if (PixelRegionsOverlap(vectorRegion, region)) { intersectsClassified = true; break; }
        }
        if (!intersectsClassified) return false;
        foreach (PdfLogicalTextBlock oldBlock in expected.TextBlocks) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeVectorWork(ref work, 1L);
            if (oldBlock.Spans.Count == 0) continue;
            PdfLogicalVisualBounds oldBounds = oldBlock.VisualBounds ?? TextBounds(expected, oldBlock);
            if (!VisualBoundsOverlap(vectorBounds, oldBounds)) continue;
            int oldSide = PaintSide(oldBlock.Spans, before.PaintOrder, ref work, cancellationToken);
            if (oldSide == 0) continue;
            PdfLogicalTextBlock? matching = null;
            bool unmatchedOppositeSide = false;
            double bestOverlap = 0D;
            foreach (PdfLogicalTextBlock newBlock in actual.TextBlocks) {
                cancellationToken.ThrowIfCancellationRequested();
                ChargeVectorWork(ref work, 1L);
                if (newBlock.Spans.Count == 0) continue;
                PdfLogicalVisualBounds newBounds = newBlock.VisualBounds ?? TextBounds(actual, newBlock);
                if (VisualBoundsOverlap(vectorBounds, newBounds)) {
                    int candidateSide = PaintSide(newBlock.Spans, after.PaintOrder, ref work, cancellationToken);
                    if (candidateSide != 0 && candidateSide != oldSide) unmatchedOppositeSide = true;
                }
                double overlap = VisualBoundsOverlapArea(oldBounds, newBounds);
                if (overlap <= bestOverlap || overlap < Math.Min(oldBounds.Width * oldBounds.Height,
                        newBounds.Width * newBounds.Height) * 0.5D) continue;
                matching = newBlock;
                bestOverlap = overlap;
            }
            if (matching != null) {
                int newSide = PaintSide(matching.Spans, after.PaintOrder, ref work, cancellationToken);
                if (newSide != 0 && newSide != oldSide) return true;
            } else if (unmatchedOppositeSide) return true;
        }
        foreach (PdfLogicalImage oldImage in expected.Images) {
            foreach (PdfImagePlacement oldPlacement in oldImage.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                ChargeVectorWork(ref work, 1L);
                if (!TryVisibleImageBounds(expected, oldPlacement, requireExactClip: true,
                        out PdfLogicalVisualBounds oldBounds) || !VisualBoundsOverlap(vectorBounds, oldBounds)) continue;
                PdfImagePlacement? matchingPlacement = null;
                double bestOverlap = 0D;
                foreach (PdfLogicalImage newImage in actual.Images) {
                    foreach (PdfImagePlacement newPlacement in newImage.Placements) {
                        cancellationToken.ThrowIfCancellationRequested();
                        ChargeVectorWork(ref work, 1L);
                        if (!TryVisibleImageBounds(actual, newPlacement, requireExactClip: true,
                                out PdfLogicalVisualBounds newBounds)) continue;
                        double overlap = VisualBoundsOverlapArea(oldBounds, newBounds);
                        if (overlap <= bestOverlap || overlap < Math.Min(oldBounds.Width * oldBounds.Height,
                                newBounds.Width * newBounds.Height) * 0.5D) continue;
                        matchingPlacement = newPlacement;
                        bestOverlap = overlap;
                    }
                }
                if (matchingPlacement != null) {
                    int oldSide = Math.Sign(oldPlacement.PaintOrder - before.PaintOrder);
                    int newSide = Math.Sign(matchingPlacement.PaintOrder - after.PaintOrder);
                    if (oldSide != 0 && newSide != 0 && oldSide != newSide) return true;
                }
            }
        }
        return false;
    }

    private static int PaintSide(IReadOnlyList<PdfTextSpan> spans, double vectorOrder,
        ref long work, CancellationToken cancellationToken) {
        int side = 0;
        foreach (PdfTextSpan span in spans) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeVectorWork(ref work, 1L);
            double order = span.PaintOrder;
            int current = order < vectorOrder ? -1 : order > vectorOrder ? 1 : 0;
            if (current == 0 || side != 0 && side != current) return 0;
            side = current;
        }
        return side;
    }

    private static bool VisualBoundsOverlap(PdfLogicalVisualBounds first, PdfLogicalVisualBounds second) =>
        VisualBoundsOverlapArea(first, second) > 0D;

    private static double VisualBoundsOverlapArea(PdfLogicalVisualBounds first, PdfLogicalVisualBounds second) =>
        Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left)) *
        Math.Max(0D, Math.Min(first.Bottom, second.Bottom) - Math.Max(first.Top, second.Top));

    private static bool PixelRegionsOverlap(PdfPixelRegion first, PdfPixelRegion second) =>
        first.X < second.X + second.Width && second.X < first.X + first.Width &&
        first.Y < second.Y + second.Height && second.Y < first.Y + first.Height;

    private static bool SameSimpleVectorPaint(PdfPageVisualPrimitive before, PdfPageVisualPrimitive after) {
        // Complex effects cannot be compared by reference across independently read PDFs.
        // Reporting uncertainty for them is safer than claiming their pixels are text or image paint.
        if (!IsSimpleVectorPaint(before) || !IsSimpleVectorPaint(after)) return false;
        return before.Kind == after.Kind && before.X == after.X && before.Y == after.Y &&
            before.Width == after.Width && before.Height == after.Height &&
            before.X1 == after.X1 && before.Y1 == after.Y1 &&
            before.X2 == after.X2 && before.Y2 == after.Y2 &&
            Nullable.Equals(before.FillColor, after.FillColor) &&
            Nullable.Equals(before.StrokeColor, after.StrokeColor) &&
            before.StrokeWidth == after.StrokeWidth &&
            before.StrokeDashStyle == after.StrokeDashStyle &&
            before.StrokeLineCap == after.StrokeLineCap &&
            before.StrokeLineJoin == after.StrokeLineJoin &&
            before.FillOpacity == after.FillOpacity && before.StrokeOpacity == after.StrokeOpacity &&
            before.FillRule == after.FillRule &&
            SameClipPath(before.ClipPath, after.ClipPath) &&
            SameDashPattern(before.StrokeDashPattern, after.StrokeDashPattern) &&
            before.PathCommands.SequenceEqual(after.PathCommands);
    }

    private static bool IsSimpleVectorPaint(PdfPageVisualPrimitive primitive) =>
        primitive.FillGradient == null && primitive.FillRadialGradient == null &&
        primitive.StrokeGradient == null && primitive.StrokeRadialGradient == null &&
        primitive.FillTilingPattern == null && primitive.StrokeTilingPattern == null;

    private static bool SameClipPath(PdfPageClipPath? before, PdfPageClipPath? after) {
        if (!before.HasValue || !after.HasValue) return before.HasValue == after.HasValue;
        PdfPageClipPath a = before.Value;
        PdfPageClipPath b = after.Value;
        return a.X == b.X && a.Y == b.Y && a.Width == b.Width && a.Height == b.Height &&
            a.IsRectangle == b.IsRectangle && a.FillRule == b.FillRule &&
            a.IsExact == b.IsExact && a.ContainsTextClipping == b.ContainsTextClipping &&
            a.Commands.SequenceEqual(b.Commands);
    }

    private static bool SameDashPattern(PdfStrokeDashPattern? before, PdfStrokeDashPattern? after) {
        if (!before.HasValue || !after.HasValue) return before.HasValue == after.HasValue;
        return before.Value.Phase == after.Value.Phase &&
            before.Value.Array.SequenceEqual(after.Value.Array);
    }

    private static bool ChangedPixelsIntersectVector(PdfPageVisualPrimitive primitive, PdfLogicalPage page,
        List<PdfPixelRegion> classified, PdfVisualPageComparison visual,
        PdfVisualComparisonOptions options, ref long work, CancellationToken cancellationToken) {
        double stroke = Math.Max(0D, primitive.StrokeWidth / 2D);
        var bounds = new PdfLogicalVisualBounds(primitive.X - stroke, primitive.Y - stroke,
            primitive.X + primitive.Width + stroke, primitive.Y + primitive.Height + stroke);
        if (!ToPixelRegion(bounds, page, visual, options, out PdfPixelRegion vector)) return false;
        foreach (PdfPixelRegion region in classified) {
            cancellationToken.ThrowIfCancellationRequested();
            ChargeVectorWork(ref work, 1L);
            int left = Math.Max(vector.X, region.X);
            int top = Math.Max(vector.Y, region.Y);
            int right = Math.Min(vector.X + vector.Width, region.X + region.Width);
            int bottom = Math.Min(vector.Y + vector.Height, region.Y + region.Height);
            if (right > left && bottom > top) {
                ChargeVectorWork(ref work, (long)(right - left) * (bottom - top));
                if (visual.HasChangedPixelsIn(new PdfPixelRegion(left, top, right - left, bottom - top), cancellationToken)) return true;
            }
        }
        return false;
    }

}
