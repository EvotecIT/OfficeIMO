using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Classifies bounded, page-local text and image evidence without claiming authoring intent.</summary>
internal static partial class PdfReviewSemanticComparer {
    internal static IReadOnlyList<PdfReviewChange> Compare(
        PdfLogicalPage expected,
        PdfLogicalPage actual,
        PdfReadPage expectedReadPage,
        PdfReadPage actualReadPage,
        PdfVisualPageComparison? visual,
        bool usesIgnoredRegions,
        PdfReviewComparisonOptions options,
        CancellationToken cancellationToken) {
        int expectedPage = expected.PageNumber;
        int actualPage = actual.PageNumber;
        var changes = new List<PdfReviewChange>();
        ValidateImagePlacementCount(expected, options, cancellationToken);
        ValidateImagePlacementCount(actual, options, cancellationToken);
        bool expectedScanned = IsScan(expected, visual, options, cancellationToken);
        bool actualScanned = IsScan(actual, visual, options, cancellationToken);
        bool scanned = expectedScanned || actualScanned;
        CompareText(expected, actual, visual, options, changes, cancellationToken);
        if (visual != null &&
            (PdfRenderCapabilities.HasIncompleteVisualProjection(visual.ExpectedCapabilityDiagnostics) ||
             PdfRenderCapabilities.HasIncompleteVisualProjection(visual.ActualCapabilityDiagnostics))) {
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.RenderUncertain,
                expectedPage, actualPage, null, null));
        }
        if (scanned) {
            if (visual is { IsMatch: false } || expectedScanned != actualScanned ||
                HasDifferentUnignoredScanImages(expected, actual, visual, options,
                    comparePayload: usesIgnoredRegions, cancellationToken: cancellationToken)) {
                changes.Add(new PdfReviewChange(PdfReviewChangeKind.ScannedPageUncertain, expectedPage, actualPage, null, null));
            }
            return OrderChanges(changes);
        }
        CompareImages(expected, actual, visual, options, changes, cancellationToken);
        if (visual is { IsMatch: false } &&
            (visual.HasSizeDifference || HasUnclassifiedPixels(expected, actual, expectedReadPage, actualReadPage,
                visual, options.Visual, changes, cancellationToken))) {
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.UnclassifiedVisual, expectedPage, actualPage, null, null));
        }
        return OrderChanges(changes);
    }

    private static void CompareText(PdfLogicalPage expected, PdfLogicalPage actual, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, List<PdfReviewChange> changes, CancellationToken cancellationToken) {
        TextFeature[] before = GetText(expected, visual, options, cancellationToken);
        TextFeature[] after = GetText(actual, visual, options, cancellationToken);
        bool[] usedBefore = new bool[before.Length];
        bool[] usedAfter = new bool[after.Length];
        MatchIdentical(before, after, static feature => feature.Normalized, usedBefore, usedAfter,
            (i, j, distance) => {
                if (!string.Equals(before[i].SemanticText, after[j].SemanticText, StringComparison.Ordinal)) {
                    changes.Add(new PdfReviewChange(PdfReviewChangeKind.TextChanged, expected.PageNumber, actual.PageNumber,
                        before[i].Bounds, after[j].Bounds, before[i].SemanticText, after[j].SemanticText,
                        canCoverRenderedPixels: before[i].Text != after[j].Text && (before[i].CanPaint || after[j].CanPaint)));
                } else if (TextGeometryDiffers(before[i], after[j], cancellationToken)) {
                    changes.Add(new PdfReviewChange(PdfReviewChangeKind.TextMoved, expected.PageNumber, actual.PageNumber,
                        before[i].Bounds, after[j].Bounds, before[i].SemanticText, after[j].SemanticText,
                        canCoverRenderedPixels: before[i].CanPaint || after[j].CanPaint));
                }
            }, cancellationToken);
        for (int i = 0; i < before.Length; i++) {
            if (usedBefore[i]) continue;
            cancellationToken.ThrowIfCancellationRequested();
            int best = FindCorresponding(before[i].Bounds, after, usedAfter);
            if (best < 0) continue;
            usedBefore[i] = true;
            usedAfter[best] = true;
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.TextChanged, expected.PageNumber, actual.PageNumber,
                        before[i].Bounds, after[best].Bounds, before[i].SemanticText, after[best].SemanticText,
                canCoverRenderedPixels: before[i].CanPaint || after[best].CanPaint));
        }
        for (int i = 0; i < before.Length; i++) if (!usedBefore[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.TextRemoved, expected.PageNumber, actual.PageNumber, before[i].Bounds, null, before[i].SemanticText,
            canCoverRenderedPixels: before[i].CanPaint));
        for (int i = 0; i < after.Length; i++) if (!usedAfter[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.TextAdded, expected.PageNumber, actual.PageNumber, null, after[i].Bounds, actualText: after[i].SemanticText,
            canCoverRenderedPixels: after[i].CanPaint));
    }

    private static void CompareImages(PdfLogicalPage expected, PdfLogicalPage actual, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, List<PdfReviewChange> changes, CancellationToken cancellationToken) {
        ImageFeature[] before = GetImages(expected, visual, options, includePayloadHash: true, cancellationToken: cancellationToken);
        ImageFeature[] after = GetImages(actual, visual, options, includePayloadHash: true, cancellationToken: cancellationToken);
        bool[] usedBefore = new bool[before.Length];
        bool[] usedAfter = new bool[after.Length];
        MatchIdentical(before, after, static feature => feature.Hash, usedBefore, usedAfter,
            (i, j, _) => {
                if (BoundsDiffer(before[i].Bounds, after[j].Bounds) || ImageTransformDiffers(before[i], after[j])) changes.Add(new PdfReviewChange(PdfReviewChangeKind.ImageMoved,
                    expected.PageNumber, actual.PageNumber, before[i].Bounds, after[j].Bounds));
            }, cancellationToken);
        for (int i = 0; i < before.Length; i++) {
            if (usedBefore[i]) continue;
            cancellationToken.ThrowIfCancellationRequested();
            int best = FindCorresponding(before[i].Bounds, after, usedAfter);
            if (best < 0) continue;
            usedBefore[i] = true;
            usedAfter[best] = true;
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.ImageChangedCandidate,
                expected.PageNumber, actual.PageNumber, before[i].Bounds, after[best].Bounds));
        }
        for (int i = 0; i < before.Length; i++) if (!usedBefore[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.ImageRemoved, expected.PageNumber, actual.PageNumber, before[i].Bounds, null));
        for (int i = 0; i < after.Length; i++) if (!usedAfter[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.ImageAdded, expected.PageNumber, actual.PageNumber, null, after[i].Bounds));
    }

    private static TextFeature[] GetText(PdfLogicalPage page, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, CancellationToken cancellationToken) {
        if (page.TextBlocks.Count > options.MaxTextBlocksPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, options.MaxTextBlocksPerPage, page.TextBlocks.Count);
        }
        var output = new List<TextFeature>(page.TextBlocks.Count);
        long remainingMappingWork = 4_000_000L;
        foreach (PdfLogicalTextBlock block in page.TextBlocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(block.Text) || block.VisualBounds is null && block.XEnd <= block.XStart) continue;
            PdfLogicalVisualBounds bounds = block.VisualBounds ?? TextBounds(page, block);
            if (IsIgnored(bounds, page, visual, options.Visual)) continue;
            string semanticText = ReconstructSemanticText(block, ref remainingMappingWork, cancellationToken);
            bool canPaint = block.Spans.Count == 0 || block.Spans.Any(span =>
                span.IsVisible && (span.Color?.A ?? 255) > 3 &&
                (!span.ClipPath.HasValue || span.CanProjectCompleteText(page) ||
                 span.ClipPath.Value is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } &&
                 !span.ClipPath.Value.CanProveNoPositiveAreaIntersection(
                     PdfPageClipPath.Rectangle(bounds.Left, bounds.Top, bounds.Width, bounds.Height))));
            output.Add(new TextFeature(block.Text, semanticText, Normalize(block.Text), bounds, canPaint, block.Spans, page));
        }
        return output.ToArray();
    }

    internal static string ReconstructSemanticText(PdfLogicalTextBlock block,
        CancellationToken cancellationToken = default) {
        long remainingMappingWork = 4_000_000L;
        return ReconstructSemanticText(block, ref remainingMappingWork, cancellationToken);
    }

    private static string ReconstructSemanticText(PdfLogicalTextBlock block,
        ref long remainingMappingWork, CancellationToken cancellationToken) {
        if (!block.Spans.Any(static span => span.HasActualText)) return block.Text;
        const long maxMappingWork = 4_000_000L;
        var replacements = new List<(int Start, int End, string Text)>(block.Spans.Count);
        IEnumerable<PdfTextSpan> spans = block.Spans;
        if (block.Spans.Count > 1 && block.Spans.All(span =>
                Math.Abs(span.Y - block.Spans[0].Y) <= Math.Max(1D, block.FontSize * 0.5D) &&
                Math.Abs(span.RotationDegrees) < 0.01D)) {
            bool rightToLeft = OfficeTextElements.ResolveBaseDirection(block.Text) == OfficeTextDirection.RightToLeft;
            spans = rightToLeft
                ? block.Spans.OrderByDescending(static span => span.X)
                : block.Spans.OrderBy(static span => span.X);
        }
        foreach (PdfTextSpan span in spans) {
            cancellationToken.ThrowIfCancellationRequested();
            if (span.Text.Length == 0) continue;
            int searchStart = 0;
            bool matched = false;
            while (searchStart < block.Text.Length) {
                long work = (long)block.Text.Length - searchStart + replacements.Count;
                if (work > remainingMappingWork) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                        maxMappingWork, maxMappingWork - remainingMappingWork + work);
                }
                remainingMappingWork -= work;
                cancellationToken.ThrowIfCancellationRequested();
                if (!PdfLogicalTextBlock.TryFindNormalizedSpan(block.Text, span.Text, searchStart, out int match, out int end)) break;
                if (!replacements.Any(item => match < item.End && end > item.Start)) {
                    replacements.Add((match, end, span.SourceActualText ?? span.Text));
                    matched = true;
                    break;
                }
                searchStart = match + 1;
            }
            if (!matched) {
                return string.Concat(spans.Select(static item => item.SourceActualText ?? item.Text));
            }
        }
        var result = new System.Text.StringBuilder(block.Text.Length);
        int offset = 0;
        foreach ((int start, int end, string text) in replacements.OrderBy(static item => item.Start)) {
            result.Append(block.Text, offset, start - offset);
            result.Append(text);
            offset = end;
        }
        result.Append(block.Text, offset, block.Text.Length - offset);
        return result.ToString();
    }

    private static ImageFeature[] GetImages(PdfLogicalPage page, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, bool includePayloadHash, CancellationToken cancellationToken) {
        int count = ValidateImagePlacementCount(page, options, cancellationToken);
        var output = new List<ImageFeature>(count);
        foreach (PdfLogicalImage image in page.Images) {
            string? hash = null;
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryVisibleImageBounds(page, placement, requireExactClip: true, out PdfLogicalVisualBounds bounds)) continue;
                if (IsIgnored(bounds, page, visual, options.Visual)) continue;
                hash ??= includePayloadHash ? Hash(image.SourceImage.EncodedBytes, cancellationToken) : string.Empty;
                output.Add(new ImageFeature(hash, bounds, VisualImageTransform(page, placement)));
            }
        }
        return output.ToArray();
    }

    private static int ValidateImagePlacementCount(PdfLogicalPage page, PdfReviewComparisonOptions options,
        CancellationToken cancellationToken) {
        int count = 0;
        foreach (PdfLogicalImage image in page.Images) {
            cancellationToken.ThrowIfCancellationRequested();
            count = checked(count + image.Placements.Count);
            if (count > options.MaxImagePlacementsPerPage) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, options.MaxImagePlacementsPerPage, count);
            }
        }
        return count;
    }

    private static bool IsScan(PdfLogicalPage page, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, CancellationToken cancellationToken) {
        (double width, double height) = page.GetVisualPageSize();
        double area = width * height;
        if (area <= 0D) return false;
        var visible = new List<PdfLogicalVisualBounds>();
        foreach (PdfLogicalImage image in page.Images) {
            if (image.SourceImage.IsImageMask || image.SourceImage.HasTransparencyMask) continue;
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (placement.Opacity < 0.999D) continue;
                // Bounds prove coverage for axis-aligned and quarter-turn placements, but not skew.
                bool axisAligned = Math.Abs(placement.B) <= 0.0001D && Math.Abs(placement.C) <= 0.0001D;
                bool quarterTurn = Math.Abs(placement.A) <= 0.0001D && Math.Abs(placement.D) <= 0.0001D &&
                    Math.Abs(placement.B) > 0.0001D && Math.Abs(placement.C) > 0.0001D;
                if (!axisAligned && !quarterTurn) continue;
                if (TryVisibleImageBounds(page, placement, requireExactClip: true, out PdfLogicalVisualBounds bounds) &&
                    !IsIgnored(bounds, page, visual, options.Visual)) visible.Add(bounds);
            }
        }
        return UnionArea(visible, cancellationToken) >= area * 0.75D;
    }

    private static bool HasDifferentUnignoredScanImages(PdfLogicalPage expected, PdfLogicalPage actual,
        PdfVisualPageComparison? visual, PdfReviewComparisonOptions options, bool comparePayload,
        CancellationToken cancellationToken) {
        ImageFeature[] before = GetImages(expected, visual, options, includePayloadHash: comparePayload, cancellationToken: cancellationToken);
        ImageFeature[] after = GetImages(actual, visual, options, includePayloadHash: comparePayload, cancellationToken: cancellationToken);
        if (before.Length != after.Length) return true;
        var matched = new bool[after.Length];
        for (int index = 0; index < before.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            bool found = false;
            for (int candidate = 0; candidate < after.Length; candidate++) {
                if (matched[candidate] ||
                    comparePayload && !string.Equals(before[index].Hash, after[candidate].Hash, StringComparison.Ordinal) ||
                    BoundsDiffer(before[index].Bounds, after[candidate].Bounds) ||
                    ImageTransformDiffers(before[index], after[candidate])) continue;
                matched[candidate] = true;
                found = true;
                break;
            }
            if (!found) return true;
        }
        return false;
    }

    private static bool HasUnclassifiedPixels(PdfLogicalPage expected, PdfLogicalPage actual,
        PdfReadPage expectedReadPage, PdfReadPage actualReadPage,
        PdfVisualPageComparison visual, PdfVisualComparisonOptions options,
        List<PdfReviewChange> changes, CancellationToken cancellationToken) {
        var classified = new List<PdfPixelRegion>(changes.Count * 2);
        foreach (PdfReviewChange change in changes) {
            if (!change.CanCoverRenderedPixels) continue;
            if (change.ExpectedBounds is PdfLogicalVisualBounds before &&
                ToPixelRegion(before, expected, visual, options, out PdfPixelRegion expectedRegion)) classified.Add(expectedRegion);
            if (change.ActualBounds is PdfLogicalVisualBounds after &&
                ToPixelRegion(after, actual, visual, options, out PdfPixelRegion actualRegion)) classified.Add(actualRegion);
        }
        if (visual.HasChangedPixelsOutside(classified, cancellationToken)) return true;
        // A text or image rectangle is only a location estimate. Where vector paint
        // intersects it, the raster change cannot be attributed to that element alone.
        return HasChangedVectorPaintInsideClassifiedBounds(expectedReadPage, actualReadPage,
            expected, actual, classified, visual, options, cancellationToken);
    }

    private static bool ToPixelRegion(PdfLogicalVisualBounds bounds, PdfLogicalPage page,
        PdfVisualPageComparison visual, PdfVisualComparisonOptions options, out PdfPixelRegion region) {
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        int rasterWidth = checked((int)Math.Ceiling(pageWidth * options.Scale));
        int rasterHeight = checked((int)Math.Ceiling(pageHeight * options.Scale));
        int offsetX = options.Alignment == PdfVisualPageAlignment.Center ? (visual.Width - rasterWidth) / 2 : 0;
        int offsetY = options.Alignment == PdfVisualPageAlignment.Center ? (visual.Height - rasterHeight) / 2 : 0;
        const int padding = 2;
        int left = ClampPixel(Math.Floor(bounds.Left * options.Scale) + offsetX - padding, visual.Width);
        int top = ClampPixel(Math.Floor(bounds.Top * options.Scale) + offsetY - padding, visual.Height);
        int right = ClampPixel(Math.Ceiling(bounds.Right * options.Scale) + offsetX + padding, visual.Width);
        int bottom = ClampPixel(Math.Ceiling(bounds.Bottom * options.Scale) + offsetY + padding, visual.Height);
        if (right <= left || bottom <= top) { region = default; return false; }
        region = new PdfPixelRegion(left, top, right - left, bottom - top);
        return true;
    }

    private static int ClampPixel(double coordinate, int extent) =>
        coordinate <= 0D ? 0 : coordinate >= extent ? extent : (int)coordinate;

    private static bool TryVisibleImageBounds(PdfLogicalPage page, PdfImagePlacement placement,
        bool requireExactClip, out PdfLogicalVisualBounds bounds) {
        bounds = default!;
        if (placement.Width <= 0D || placement.Height <= 0D || placement.IsHiddenOptionalContent ||
            placement.Opacity <= 0D || placement.HasSoftMask || placement.HasUnsupportedImagePaintEffect) return false;
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
            placement.X + placement.Width, placement.Y + placement.Height);
        double left = Math.Max(0D, mapped.Left);
        double top = Math.Max(0D, mapped.Top);
        double right = Math.Min(pageWidth, mapped.Right);
        double bottom = Math.Min(pageHeight, mapped.Bottom);
        if (placement.Clip is { } clip) {
            if (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping) {
                if (requireExactClip) return false;
            } else {
                PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                    page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
                left = Math.Max(left, clipped.Left);
                top = Math.Max(top, clipped.Top);
                right = Math.Min(right, clipped.Right);
                bottom = Math.Min(bottom, clipped.Bottom);
            }
        }
        if (right <= left || bottom <= top) return false;
        bounds = new PdfLogicalVisualBounds(left, top, right, bottom);
        return true;
    }

    private static double UnionArea(List<PdfLogicalVisualBounds> bounds, CancellationToken cancellationToken) {
        if (bounds.Count == 0) return 0D;
        double[] edges = bounds.SelectMany(static item => new[] { item.Left, item.Right }).Distinct().OrderBy(static x => x).ToArray();
        double area = 0D;
        for (int edge = 0; edge < edges.Length - 1; edge++) {
            cancellationToken.ThrowIfCancellationRequested();
            double left = edges[edge];
            double right = edges[edge + 1];
            if (right <= left) continue;
            var spans = bounds.Where(item => item.Left < right && item.Right > left)
                .OrderBy(static item => item.Top).ToArray();
            double covered = 0D;
            double top = 0D;
            double bottom = 0D;
            bool hasSpan = false;
            foreach (PdfLogicalVisualBounds span in spans) {
                if (!hasSpan) {
                    top = span.Top;
                    bottom = span.Bottom;
                    hasSpan = true;
                } else if (span.Top <= bottom) {
                    bottom = Math.Max(bottom, span.Bottom);
                } else {
                    covered += bottom - top;
                    top = span.Top;
                    bottom = span.Bottom;
                }
            }
            if (hasSpan) covered += bottom - top;
            area += (right - left) * covered;
        }
        return area;
    }

    private static PdfLogicalVisualBounds TextBounds(PdfLogicalPage page, PdfLogicalTextBlock block) {
        double size = Math.Max(1D, block.FontSize);
        PdfVisualBounds mapped = page.TransformBoundsToVisual(block.XStart, block.BaselineY - size * 0.25D, block.XEnd, block.BaselineY + size * 0.85D);
        return new PdfLogicalVisualBounds(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
    }

    private static bool IsIgnored(PdfLogicalVisualBounds bounds, PdfLogicalPage page, PdfVisualPageComparison? visual,
        PdfVisualComparisonOptions options) {
        if (options.IgnoredRegions.Count == 0) return false;
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        int rasterWidth = checked((int)Math.Ceiling(pageWidth * options.Scale));
        int rasterHeight = checked((int)Math.Ceiling(pageHeight * options.Scale));
        int offsetX = options.Alignment == PdfVisualPageAlignment.Center ? ((visual?.Width ?? rasterWidth) - rasterWidth) / 2 : 0;
        int offsetY = options.Alignment == PdfVisualPageAlignment.Center ? ((visual?.Height ?? rasterHeight) - rasterHeight) / 2 : 0;
        double left = bounds.Left * options.Scale + offsetX;
        double top = bounds.Top * options.Scale + offsetY;
        double right = bounds.Right * options.Scale + offsetX;
        double bottom = bounds.Bottom * options.Scale + offsetY;
        return options.IgnoredRegions.Any(region => region.X <= left && region.Y <= top &&
            (long)region.X + region.Width >= right && (long)region.Y + region.Height >= bottom);
    }

    private static void MatchIdentical<T>(T[] before, T[] after, Func<T, string> key,
        bool[] usedBefore, bool[] usedAfter, Action<int, int, double> onMatch, CancellationToken cancellationToken)
        where T : IBoundedFeature {
        // Claim exact surviving instances first, so removing one of several identical items does not look like a move.
        var candidates = new List<(int Index, double Distance)>();
        for (int i = 0; i < before.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            double nearest = double.MaxValue;
            for (int j = 0; j < after.Length; j++) {
                if (!string.Equals(key(before[i]), key(after[j]), StringComparison.Ordinal)) continue;
                nearest = Math.Min(nearest, Distance(before[i].Bounds, after[j].Bounds));
            }
            if (nearest < double.MaxValue) candidates.Add((i, nearest));
        }
        foreach ((int i, _) in candidates.OrderBy(static candidate => candidate.Distance).ThenBy(static candidate => candidate.Index)) {
            cancellationToken.ThrowIfCancellationRequested();
            int best = -1;
            double distance = double.MaxValue;
            for (int j = 0; j < after.Length; j++) {
                if (usedAfter[j] || !string.Equals(key(before[i]), key(after[j]), StringComparison.Ordinal)) continue;
                double candidateDistance = Distance(before[i].Bounds, after[j].Bounds);
                if (candidateDistance < distance) { best = j; distance = candidateDistance; }
            }
            if (best < 0) continue;
            usedBefore[i] = true;
            usedAfter[best] = true;
            onMatch(i, best, distance);
        }
    }

    private static int FindCorresponding<T>(PdfLogicalVisualBounds bounds, T[] other, bool[] used) where T : IBoundedFeature {
        int best = -1;
        double score = double.MaxValue;
        for (int i = 0; i < other.Length; i++) {
            if (used[i]) continue;
            PdfLogicalVisualBounds candidate = other[i].Bounds;
            double overlap = Overlap(bounds, candidate);
            double distance = Distance(bounds, candidate);
            if (overlap < 0.15D && distance > 20D) continue;
            if (distance < score) { best = i; score = distance; }
        }
        return best;
    }

    private static double Overlap(PdfLogicalVisualBounds first, PdfLogicalVisualBounds second) {
        double area = Math.Max(0D, Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left)) *
            Math.Max(0D, Math.Min(first.Bottom, second.Bottom) - Math.Max(first.Top, second.Top));
        double smaller = Math.Min((first.Right - first.Left) * (first.Bottom - first.Top),
            (second.Right - second.Left) * (second.Bottom - second.Top));
        return smaller <= 0D ? 0D : area / smaller;
    }

    private static double Distance(PdfLogicalVisualBounds first, PdfLogicalVisualBounds second) {
        double dx = (first.Left + first.Right - second.Left - second.Right) / 2D;
        double dy = (first.Top + first.Bottom - second.Top - second.Bottom) / 2D;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private static bool BoundsDiffer(PdfLogicalVisualBounds first, PdfLogicalVisualBounds second) =>
        Math.Abs(first.Left - second.Left) > 0.01D || Math.Abs(first.Top - second.Top) > 0.01D ||
        Math.Abs(first.Right - second.Right) > 0.01D || Math.Abs(first.Bottom - second.Bottom) > 0.01D;

    private static bool TextGeometryDiffers(TextFeature first, TextFeature second, CancellationToken cancellationToken) {
        if (BoundsDiffer(first.Bounds, second.Bounds)) return true;
        var before = new List<GlyphGeometry>();
        var after = new List<GlyphGeometry>();
        if (TryGetGlyphGeometry(first, before, cancellationToken) &&
            TryGetGlyphGeometry(second, after, cancellationToken)) {
            if (before.Count != after.Count) return true;
            for (int i = 0; i < before.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (before[i].DiffersFrom(after[i])) return true;
            }
            return false;
        }
        // Without usable advances, different segmentation cannot prove equivalent placement.
        if (first.Spans.Count != second.Spans.Count) return true;
        for (int i = 0; i < first.Spans.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfTextSpan left = first.Spans[i];
            PdfTextSpan right = second.Spans[i];
            PdfTextSpanBounds leftBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(left);
            PdfTextSpanBounds rightBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(right);
            PdfVisualBounds leftVisual = first.Page.TransformBoundsToVisual(leftBounds.Left, leftBounds.Bottom,
                leftBounds.Right, leftBounds.Top);
            PdfVisualBounds rightVisual = second.Page.TransformBoundsToVisual(rightBounds.Left, rightBounds.Bottom,
                rightBounds.Right, rightBounds.Top);
            if (Math.Abs(leftVisual.Left - rightVisual.Left) > 0.01D ||
                Math.Abs(leftVisual.Top - rightVisual.Top) > 0.01D ||
                Math.Abs(leftVisual.Right - rightVisual.Right) > 0.01D ||
                Math.Abs(leftVisual.Bottom - rightVisual.Bottom) > 0.01D ||
                TextDirectionDiffers(first.Page, left, second.Page, right)) return true;
        }
        return false;
    }

    private static bool TryGetGlyphGeometry(TextFeature feature, List<GlyphGeometry> output,
        CancellationToken cancellationToken) {
        const int maxGlyphsPerBlock = 200_000;
        foreach (PdfTextSpan span in feature.Spans) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, cancellationToken, out double[] boundaries)) {
                if (span.Advance <= 0D || span.Text.Length == 0) return false;
                boundaries = new double[span.Text.Length + 1];
                for (int i = 1; i < boundaries.Length; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    boundaries[i] = span.Advance * i / span.Text.Length;
                }
            }
            if (span.Text.Length > maxGlyphsPerBlock - output.Count) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                    maxGlyphsPerBlock, (long)output.Count + span.Text.Length);
            }
            (double x0, double y0) = VisualPoint(feature.Page, span.X, span.Y);
            double radians = span.RotationDegrees * Math.PI / 180D;
            (double x1, double y1) = VisualPoint(feature.Page,
                span.X + Math.Cos(radians), span.Y + Math.Sin(radians));
            for (int i = 0; i < span.Text.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                double start = Math.Min(boundaries[i], boundaries[i + 1]);
                double advance = Math.Abs(boundaries[i + 1] - boundaries[i]);
                PdfTextSpanBounds glyph = PdfTextSpanGeometry.GetAxisAlignedBounds(span, start, advance);
                PdfVisualBounds visual = feature.Page.TransformBoundsToVisual(glyph.Left, glyph.Bottom,
                    glyph.Right, glyph.Top);
                output.Add(new GlyphGeometry(span.Text[i], visual, x1 - x0, y1 - y0));
            }
        }
        return true;
    }

    private static bool TextDirectionDiffers(PdfLogicalPage firstPage, PdfTextSpan first,
        PdfLogicalPage secondPage, PdfTextSpan second) {
        double a = first.RotationDegrees * Math.PI / 180D;
        double b = second.RotationDegrees * Math.PI / 180D;
        (double ax0, double ay0) = VisualPoint(firstPage, first.X, first.Y);
        (double ax1, double ay1) = VisualPoint(firstPage, first.X + Math.Cos(a), first.Y + Math.Sin(a));
        (double bx0, double by0) = VisualPoint(secondPage, second.X, second.Y);
        (double bx1, double by1) = VisualPoint(secondPage, second.X + Math.Cos(b), second.Y + Math.Sin(b));
        return Math.Abs((ax1 - ax0) - (bx1 - bx0)) > 0.01D ||
            Math.Abs((ay1 - ay0) - (by1 - by0)) > 0.01D;
    }

    private readonly struct GlyphGeometry {
        internal GlyphGeometry(char character, PdfVisualBounds bounds, double directionX, double directionY) {
            Character = character; Bounds = bounds; DirectionX = directionX; DirectionY = directionY;
        }
        private char Character { get; }
        private PdfVisualBounds Bounds { get; }
        private double DirectionX { get; }
        private double DirectionY { get; }
        internal bool DiffersFrom(GlyphGeometry other) => Character != other.Character ||
            Math.Abs(Bounds.Left - other.Bounds.Left) > 0.01D ||
            Math.Abs(Bounds.Top - other.Bounds.Top) > 0.01D ||
            Math.Abs(Bounds.Right - other.Bounds.Right) > 0.01D ||
            Math.Abs(Bounds.Bottom - other.Bounds.Bottom) > 0.01D ||
            Math.Abs(DirectionX - other.DirectionX) > 0.01D ||
            Math.Abs(DirectionY - other.DirectionY) > 0.01D;
    }

    private static Matrix2D VisualImageTransform(PdfLogicalPage page, PdfImagePlacement placement) {
        (double x0, double y0) = VisualPoint(page, placement.E, placement.F);
        (double x1, double y1) = VisualPoint(page, placement.A + placement.E, placement.B + placement.F);
        (double x2, double y2) = VisualPoint(page, placement.C + placement.E, placement.D + placement.F);
        return new Matrix2D(x1 - x0, y1 - y0, x2 - x0, y2 - y0, x0, y0);
    }

    private static (double X, double Y) VisualPoint(PdfLogicalPage page, double x, double y) {
        PdfVisualBounds point = page.TransformBoundsToVisual(x, y, x, y);
        return (point.Left, point.Top);
    }

    private static bool ImageTransformDiffers(ImageFeature first, ImageFeature second) =>
        Math.Abs(first.VisualTransform.A - second.VisualTransform.A) > 0.01D ||
        Math.Abs(first.VisualTransform.B - second.VisualTransform.B) > 0.01D ||
        Math.Abs(first.VisualTransform.C - second.VisualTransform.C) > 0.01D ||
        Math.Abs(first.VisualTransform.D - second.VisualTransform.D) > 0.01D ||
        Math.Abs(first.VisualTransform.E - second.VisualTransform.E) > 0.01D ||
        Math.Abs(first.VisualTransform.F - second.VisualTransform.F) > 0.01D;

    private static string Normalize(string text) => string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static PdfReviewChange[] OrderChanges(List<PdfReviewChange> changes) =>
        changes.OrderBy(static change => change.ExpectedBounds?.Top ?? change.ActualBounds?.Top ?? double.MaxValue)
            .ThenBy(static change => change.ExpectedBounds?.Left ?? change.ActualBounds?.Left ?? double.MaxValue)
            .ToArray();

    private static string Hash(byte[] bytes, CancellationToken cancellationToken) {
        using (SHA256 sha = SHA256.Create()) {
            const int chunkSize = 64 * 1024;
            int offset = 0;
            while (offset < bytes.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(chunkSize, bytes.Length - offset);
                sha.TransformBlock(bytes, offset, count, bytes, offset);
                offset += count;
            }
            sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return Convert.ToBase64String(sha.Hash!);
        }
    }

    private interface IBoundedFeature { PdfLogicalVisualBounds Bounds { get; } }
    private sealed class TextFeature : IBoundedFeature {
        internal TextFeature(string text, string semanticText, string normalized, PdfLogicalVisualBounds bounds, bool canPaint,
            IReadOnlyList<PdfTextSpan> spans, PdfLogicalPage page) {
            Text = text;
            SemanticText = semanticText;
            Normalized = normalized;
            Bounds = bounds;
            CanPaint = canPaint;
            Spans = spans;
            Page = page;
        }
        internal string Text { get; }
        internal string SemanticText { get; }
        internal string Normalized { get; }
        internal bool CanPaint { get; }
        internal IReadOnlyList<PdfTextSpan> Spans { get; }
        internal PdfLogicalPage Page { get; }
        public PdfLogicalVisualBounds Bounds { get; }
    }
    private sealed class ImageFeature : IBoundedFeature {
        internal ImageFeature(string hash, PdfLogicalVisualBounds bounds, Matrix2D visualTransform) {
            Hash = hash; Bounds = bounds; VisualTransform = visualTransform;
        }
        internal string Hash { get; }
        internal Matrix2D VisualTransform { get; }
        public PdfLogicalVisualBounds Bounds { get; }
    }
}
