namespace OfficeIMO.Pdf;

/// <summary>
/// Builds one semantic image region per visible placement and associates captions from geometry and paint order.
/// </summary>
internal sealed class PdfImageRegionDetectionStage : IPdfImageRegionDetectionStage {
    private const double CoordinateTolerance = 1D;

    public IReadOnlyList<PdfUnderstandingImageRegion> Detect(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingSemanticElement> semanticElements) {
        Guard.NotNull(context, nameof(context));
        Guard.NotNull(semanticElements, nameof(semanticElements));
        if (context.ImagePlacements.Count == 0) return Array.Empty<PdfUnderstandingImageRegion>();

        var candidates = new List<CaptionAssociationCandidate>();
        for (int imageIndex = 0; imageIndex < context.ImagePlacements.Count; imageIndex++) {
            PdfImagePlacement placement = context.ImagePlacements[imageIndex];
            for (int semanticIndex = 0; semanticIndex < semanticElements.Count; semanticIndex++) {
                context.ConsumeWork();
                PdfUnderstandingSemanticElement semantic = semanticElements[semanticIndex];
                if (semantic.Kind != PdfUnderstandingSemanticKind.Caption ||
                    semantic.Evidence.Any(static evidence => string.Equals(
                        evidence.Code,
                        "semantic.table-caption-geometry",
                        StringComparison.Ordinal))) continue;
                if (!TryScoreCaption(context, semantic.Region, placement, out double score, out double gap, out double alignment)) continue;
                if (candidates.Count >= context.MaxImageCaptionCandidatesPerPage) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.UnderstandingArtifacts,
                        context.MaxImageCaptionCandidatesPerPage,
                        candidates.Count + 1L);
                }
                double? paintDistance = GetPaintDistance(semantic.Region, placement);
                candidates.Add(new CaptionAssociationCandidate(
                    imageIndex,
                    semanticIndex,
                    score,
                    gap,
                    alignment,
                    paintDistance));
            }
        }

        CaptionAssociationCandidate[] orderedCandidates = PdfAdvancedUnderstandingStages.CopyAndSort(
            context,
            candidates,
            static (left, right) => CompareCandidates(left, right));
        var captionsByImage = new Dictionary<int, CaptionAssociationCandidate>();
        var consumedCaptions = new HashSet<int>();
        for (int index = 0; index < orderedCandidates.Length; index++) {
            context.ConsumeWork();
            CaptionAssociationCandidate candidate = orderedCandidates[index];
            if (captionsByImage.ContainsKey(candidate.ImageIndex) || !consumedCaptions.Add(candidate.SemanticIndex)) continue;
            captionsByImage.Add(candidate.ImageIndex, candidate);
        }

        var result = new PdfUnderstandingImageRegion[context.ImagePlacements.Count];
        for (int imageIndex = 0; imageIndex < result.Length; imageIndex++) {
            context.ConsumeWork();
            PdfImagePlacement placement = context.ImagePlacements[imageIndex];
            var evidence = new List<PdfInferenceEvidence> {
                new PdfInferenceEvidence(
                    "image-region.placement-geometry",
                    "The content stream provides an image placement invocation and page-space bounds.",
                    0.75D)
            };
            if (placement.MarkedContentId.HasValue) {
                evidence.Add(new PdfInferenceEvidence(
                    "image-region.marked-content",
                    "The image placement preserves tagged-PDF marked-content ownership.",
                    0.95D));
            }
            PdfUnderstandingSemanticElement? caption = null;
            double confidence = placement.MarkedContentId.HasValue
                ? 0.9D
                : HasUsableGeometry(placement) ? 0.75D : 0.5D;
            if (captionsByImage.TryGetValue(imageIndex, out CaptionAssociationCandidate association)) {
                caption = semanticElements[association.SemanticIndex];
                confidence = PdfInference.Clamp(0.78D + association.Score * 0.18D);
                evidence.Add(new PdfInferenceEvidence(
                    "image-region.caption-proximity",
                    "A caption region is vertically adjacent and horizontally aligned with the image placement.",
                    association.Score));
                if (association.PaintDistance.HasValue) {
                    evidence.Add(new PdfInferenceEvidence(
                        "image-region.paint-order",
                        "Image and caption paint positions provide a deterministic association tie-break.",
                        Math.Max(0.1D, 1D / (1D + association.PaintDistance.Value))));
                }
            }
            result[imageIndex] = new PdfUnderstandingImageRegion(placement, caption, confidence, evidence);
        }
        return Array.AsReadOnly(result);
    }

    internal static bool IsAdjacentImageCaption(
        PdfUnderstandingPageContext context,
        PdfUnderstandingRegion region) {
        if (!IsEligibleCaptionRegion(region)) return false;
        for (int imageIndex = 0; imageIndex < context.ImagePlacements.Count; imageIndex++) {
            context.ConsumeWork();
            if (TryScoreCaption(context, region, context.ImagePlacements[imageIndex], out _, out _, out _)) return true;
        }
        return false;
    }

    private static bool TryScoreCaption(
        PdfUnderstandingPageContext context,
        PdfUnderstandingRegion region,
        PdfImagePlacement placement,
        out double score,
        out double gap,
        out double alignment) {
        score = 0D;
        gap = double.MaxValue;
        alignment = 0D;
        if (!IsEligibleCaptionRegion(region) || !TryGetVisualBounds(context, region, out PdfVisualBounds captionBounds)) return false;
        if (!HasUsableGeometry(placement) ||
            !PdfPageInteractionMap.TryGetVisibleImageBounds(context.Page, placement, out PdfVisualBounds imageBounds)) return false;
        double imageWidth = imageBounds.Right - imageBounds.Left;
        double imageHeight = imageBounds.Bottom - imageBounds.Top;
        double captionWidth = captionBounds.Right - captionBounds.Left;
        double fontSize = Math.Max(1D, region.Lines.Max(static line => line.FontSize));
        if (imageWidth <= 0D || imageHeight <= 0D || captionWidth <= 0D) return false;
        if (Math.Max(imageWidth, imageHeight) < fontSize * 4D || imageWidth * imageHeight < fontSize * fontSize * 24D) return false;
        if (captionWidth > imageWidth * 1.6D) return false;

        double aboveGap = imageBounds.Top - captionBounds.Bottom;
        double belowGap = captionBounds.Top - imageBounds.Bottom;
        if (aboveGap >= -CoordinateTolerance) gap = Math.Max(0D, aboveGap);
        if (belowGap >= -CoordinateTolerance) gap = Math.Min(gap, Math.Max(0D, belowGap));
        double maximumGap = Math.Max(10D, fontSize * 2.75D);
        if (gap > maximumGap) return false;

        double overlap = Math.Max(0D, Math.Min(imageBounds.Right, captionBounds.Right) - Math.Max(imageBounds.Left, captionBounds.Left));
        double overlapRatio = overlap / Math.Min(imageWidth, captionWidth);
        double centerDistance = Math.Abs(
            (imageBounds.Left + imageBounds.Right) / 2D -
            (captionBounds.Left + captionBounds.Right) / 2D);
        double centerAlignment = Math.Max(0D, 1D - centerDistance / Math.Max(imageWidth, captionWidth));
        alignment = Math.Max(overlapRatio, centerAlignment);
        if (overlapRatio < 0.55D && centerAlignment < 0.78D) return false;

        double gapScore = Math.Max(0D, 1D - gap / maximumGap);
        score = PdfInference.Clamp(gapScore * 0.58D + alignment * 0.42D);
        return score >= 0.62D;
    }

    private static bool IsEligibleCaptionRegion(PdfUnderstandingRegion region) {
        string text = region.Text.Trim();
        return region.Lines.Count is >= 1 and <= 3 &&
            text.Length > 0 &&
            PdfUnicodeScalarAnalysis.CountScalars(text) <= 512 &&
            !region.Evidence.Any(static evidence => string.Equals(
                evidence.Code,
                "region.canonical-table",
                StringComparison.Ordinal));
    }

    private static bool TryGetVisualBounds(
        PdfUnderstandingPageContext context,
        PdfUnderstandingRegion region,
        out PdfVisualBounds bounds) {
        PdfLogicalVisualBounds[] direct = region.Lines
            .Select(static line => line.VisualBounds)
            .Where(static value => value is not null)
            .Cast<PdfLogicalVisualBounds>()
            .ToArray();
        if (direct.Length == region.Lines.Count) {
            bounds = new PdfVisualBounds(
                direct.Min(static value => value.Left),
                direct.Min(static value => value.Top),
                direct.Max(static value => value.Right),
                direct.Max(static value => value.Bottom));
            return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
        }

        double largestFontSize = Math.Max(1D, region.Lines.Max(static line => line.FontSize));
        double bottom = region.YBottom - largestFontSize * 0.25D;
        double top = region.YTop + largestFontSize;
        bounds = context.Page.TransformBoundsToVisual(region.XStart, bottom, region.XEnd, top);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static bool HasUsableGeometry(PdfImagePlacement placement) =>
        IsFinite(placement.X) &&
        IsFinite(placement.Y) &&
        IsFinite(placement.Width) &&
        IsFinite(placement.Height) &&
        placement.Width > 0D &&
        placement.Height > 0D;

    private static double? GetPaintDistance(
        PdfUnderstandingRegion region,
        PdfImagePlacement placement) {
        double[] paintOrders = region.Lines
            .SelectMany(static line => line.Words)
            .SelectMany(static word => word.SourceRuns)
            .Select(static run => run.PaintOrder)
            .Where(static value => IsFinite(value))
            .ToArray();
        if (paintOrders.Length == 0 || !IsFinite(placement.PaintOrder)) return null;
        double nearest = paintOrders.Min(value => Math.Abs(value - placement.PaintOrder));
        return IsFinite(nearest) ? nearest : null;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static int CompareCandidates(CaptionAssociationCandidate left, CaptionAssociationCandidate right) {
        int score = right.Score.CompareTo(left.Score);
        if (score != 0) return score;
        int gap = left.Gap.CompareTo(right.Gap);
        if (gap != 0) return gap;
        int alignment = right.Alignment.CompareTo(left.Alignment);
        if (alignment != 0) return alignment;
        int paint = ComparePaintDistance(left.PaintDistance, right.PaintDistance);
        if (paint != 0) return paint;
        int image = left.ImageIndex.CompareTo(right.ImageIndex);
        return image != 0 ? image : left.SemanticIndex.CompareTo(right.SemanticIndex);
    }

    private static int ComparePaintDistance(double? left, double? right) {
        if (left.HasValue && right.HasValue) return left.Value.CompareTo(right.Value);
        if (left.HasValue) return -1;
        if (right.HasValue) return 1;
        return 0;
    }

    private readonly record struct CaptionAssociationCandidate(
        int ImageIndex,
        int SemanticIndex,
        double Score,
        double Gap,
        double Alignment,
        double? PaintDistance);
}
