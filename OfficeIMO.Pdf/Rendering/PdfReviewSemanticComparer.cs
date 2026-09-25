using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Classifies bounded, page-local text and image evidence without claiming authoring intent.</summary>
internal static class PdfReviewSemanticComparer {
    internal static IReadOnlyList<PdfReviewChange> Compare(
        PdfLogicalPage expected,
        PdfLogicalPage actual,
        PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options,
        CancellationToken cancellationToken) {
        int expectedPage = expected.PageNumber;
        int actualPage = actual.PageNumber;
        var changes = new List<PdfReviewChange>();
        ValidateImagePlacementCount(expected, options, cancellationToken);
        ValidateImagePlacementCount(actual, options, cancellationToken);
        bool scanned = IsScan(expected, cancellationToken) || IsScan(actual, cancellationToken);
        CompareText(expected, actual, visual, options, changes, cancellationToken);
        if (scanned) {
            if (visual is { IsMatch: false }) changes.Add(new PdfReviewChange(PdfReviewChangeKind.ScannedPageUncertain, expectedPage, actualPage, null, null));
            return OrderChanges(changes);
        }
        CompareImages(expected, actual, visual, options, changes, cancellationToken);
        if (changes.Count == 0 && visual is { IsMatch: false }) {
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.UnclassifiedVisual, expectedPage, actualPage, null, null));
        }
        return OrderChanges(changes);
    }

    private static void CompareText(PdfLogicalPage expected, PdfLogicalPage actual, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, List<PdfReviewChange> changes, CancellationToken cancellationToken) {
        TextFeature[] before = GetText(expected, visual, options);
        TextFeature[] after = GetText(actual, visual, options);
        bool[] usedBefore = new bool[before.Length];
        bool[] usedAfter = new bool[after.Length];
        MatchIdentical(before, after, static feature => feature.Normalized, usedBefore, usedAfter,
            (i, j, distance) => {
                if (distance > 4D) changes.Add(new PdfReviewChange(PdfReviewChangeKind.TextMoved, expected.PageNumber, actual.PageNumber,
                    before[i].Bounds, after[j].Bounds, before[i].Text, after[j].Text));
            }, cancellationToken);
        for (int i = 0; i < before.Length; i++) {
            if (usedBefore[i]) continue;
            cancellationToken.ThrowIfCancellationRequested();
            int best = FindCorresponding(before[i].Bounds, after, usedAfter);
            if (best < 0) continue;
            usedBefore[i] = true;
            usedAfter[best] = true;
            changes.Add(new PdfReviewChange(PdfReviewChangeKind.TextChanged, expected.PageNumber, actual.PageNumber,
                before[i].Bounds, after[best].Bounds, before[i].Text, after[best].Text));
        }
        for (int i = 0; i < before.Length; i++) if (!usedBefore[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.TextRemoved, expected.PageNumber, actual.PageNumber, before[i].Bounds, null, before[i].Text));
        for (int i = 0; i < after.Length; i++) if (!usedAfter[i]) changes.Add(new PdfReviewChange(
            PdfReviewChangeKind.TextAdded, expected.PageNumber, actual.PageNumber, null, after[i].Bounds, actualText: after[i].Text));
    }

    private static void CompareImages(PdfLogicalPage expected, PdfLogicalPage actual, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, List<PdfReviewChange> changes, CancellationToken cancellationToken) {
        ImageFeature[] before = GetImages(expected, visual, options, cancellationToken);
        ImageFeature[] after = GetImages(actual, visual, options, cancellationToken);
        bool[] usedBefore = new bool[before.Length];
        bool[] usedAfter = new bool[after.Length];
        MatchIdentical(before, after, static feature => feature.Hash, usedBefore, usedAfter,
            (i, j, distance) => {
                if (distance > 4D) changes.Add(new PdfReviewChange(PdfReviewChangeKind.ImageMoved,
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

    private static TextFeature[] GetText(PdfLogicalPage page, PdfVisualPageComparison? visual, PdfReviewComparisonOptions options) {
        if (page.TextBlocks.Count > options.MaxTextBlocksPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, options.MaxTextBlocksPerPage, page.TextBlocks.Count);
        }
        var output = new List<TextFeature>(page.TextBlocks.Count);
        foreach (PdfLogicalTextBlock block in page.TextBlocks) {
            if (string.IsNullOrWhiteSpace(block.Text) || block.VisualBounds is null && block.XEnd <= block.XStart) continue;
            PdfLogicalVisualBounds bounds = block.VisualBounds ?? TextBounds(page, block);
            if (IsIgnored(bounds, page, visual, options.Visual)) continue;
            output.Add(new TextFeature(block.Text, Normalize(block.Text), bounds));
        }
        return output.ToArray();
    }

    private static ImageFeature[] GetImages(PdfLogicalPage page, PdfVisualPageComparison? visual,
        PdfReviewComparisonOptions options, CancellationToken cancellationToken) {
        int count = ValidateImagePlacementCount(page, options, cancellationToken);
        var output = new List<ImageFeature>(count);
        foreach (PdfLogicalImage image in page.Images) {
            string? hash = null;
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (placement.Width <= 0D || placement.Height <= 0D || placement.IsHiddenOptionalContent) continue;
                PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y, placement.X + placement.Width, placement.Y + placement.Height);
                var bounds = new PdfLogicalVisualBounds(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
                if (IsIgnored(bounds, page, visual, options.Visual)) continue;
                hash ??= Hash(image.SourceImage.EncodedBytes, cancellationToken);
                output.Add(new ImageFeature(hash, bounds));
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

    private static bool IsScan(PdfLogicalPage page, CancellationToken cancellationToken) {
        (double width, double height) = page.GetVisualPageSize();
        double area = width * height;
        if (area <= 0D) return false;
        foreach (PdfLogicalImage image in page.Images) {
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (placement.Width <= 0D || placement.Height <= 0D || placement.IsHiddenOptionalContent) continue;
                PdfVisualBounds bounds = page.TransformBoundsToVisual(placement.X, placement.Y, placement.X + placement.Width, placement.Y + placement.Height);
                double left = Math.Max(0D, bounds.Left);
                double top = Math.Max(0D, bounds.Top);
                double right = Math.Min(width, bounds.Right);
                double bottom = Math.Min(height, bounds.Bottom);
                if (placement.Clip is { } clip) {
                    if (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping) continue;
                    PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                        page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
                    left = Math.Max(left, clipped.Left);
                    top = Math.Max(top, clipped.Top);
                    right = Math.Min(right, clipped.Right);
                    bottom = Math.Min(bottom, clipped.Bottom);
                }
                if (Math.Max(0D, right - left) * Math.Max(0D, bottom - top) >= area * 0.75D) return true;
            }
        }
        return false;
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
        internal TextFeature(string text, string normalized, PdfLogicalVisualBounds bounds) { Text = text; Normalized = normalized; Bounds = bounds; }
        internal string Text { get; }
        internal string Normalized { get; }
        public PdfLogicalVisualBounds Bounds { get; }
    }
    private sealed class ImageFeature : IBoundedFeature {
        internal ImageFeature(string hash, PdfLogicalVisualBounds bounds) { Hash = hash; Bounds = bounds; }
        internal string Hash { get; }
        public PdfLogicalVisualBounds Bounds { get; }
    }
}
