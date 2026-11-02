using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Shared by every layout context and pass contributing to one generated PDF.
    // Account for attempted copies as well as retained annotations: clipping a
    // result does not refund the allocation and transformation work already done.
    private sealed class DrawingPatternAnnotationBudget {
        private long used;
        private int limit;

        internal void Observe(int maximumTileCount) => limit = Math.Max(limit, maximumTileCount);

        internal void Reserve(int tileCount, int linksPerTile) {
            long work = (long)tileCount * linksPerTile;
            if (work > limit - used)
                throw new InvalidOperationException("PDF pattern annotation aggregate expansion exceeds the configured tile-count limit.");
            used += work;
        }
    }

    private sealed partial class LayoutContext {
        private readonly DrawingPatternAnnotationBudget drawingPatternAnnotations;
        private void DrawDrawingPatternAt(OfficeDrawingTilingPattern pattern, double originX, double originTopY, OfficeDrawingTextMetrics metrics) {
            cancellationToken.ThrowIfCancellationRequested();
            if (pattern.Opacity <= 0D) return;
            drawingPatternAnnotations.Observe(pattern.MaximumTileCount);
            bool previousWrappers = _suppressCanvasAccessibilityWrappers;
            bool previousText = _suppressCanvasActualTextChildren;
            OfficeTransform previousEffectToPage = _canvasEffectToPage;
            var previousSourceStructure = _drawingSourceStructure;
            int annotationStart = currentPage!.Annotations.Count;
            _drawingSourceStructure = null;
            _suppressCanvasAccessibilityWrappers = true;
            _suppressCanvasActualTextChildren = true;
            // Tile-local bounds must survive nested transparency Forms even when
            // the reusable tile is larger than the destination page.
            _canvasEffectToPage = OfficeTransform.Scale(currentOpts.PageWidth.ToPoints() / pattern.InnerTile.Width,
                currentOpts.PageHeight.ToPoints() / pattern.InnerTile.Height);
            try {
                RenderEffectGroup(OfficeTransform.Identity, pattern.Opacity, OfficeBlendMode.Normal, () => {
                    DrawDrawingElements(pattern.InnerTile, 0D, pattern.InnerTile.Height, metrics);
                    ClipCanvasLinkAnnotations(currentPage.Annotations, annotationStart, 0D, 0D,
                        pattern.InnerTile.Width, pattern.InnerTile.Height,
                        OfficeClipPath.Rectangle(pattern.InnerTile.Width, pattern.InnerTile.Height));
                }, forceForm: true);
            } finally {
                _drawingSourceStructure = previousSourceStructure;
                _suppressCanvasAccessibilityWrappers = previousWrappers;
                _suppressCanvasActualTextChildren = previousText;
                _canvasEffectToPage = previousEffectToPage;
            }
            PageEffectGroup effect = currentPage.EffectGroups[currentPage.EffectGroups.Count - 1];
            effect.TilingPattern = new PageTilingPattern(pattern, originX, originTopY);
            effect.BoundsLeft = originX + pattern.Area.X;
            effect.BoundsBottom = originTopY - pattern.Area.Y - pattern.Area.Height;
            effect.BoundsRight = effect.BoundsLeft + pattern.Area.Width;
            effect.BoundsTop = effect.BoundsBottom + pattern.Area.Height;

            // Paint is stored once. Only interactive annotations need expansion,
            // because PDF pattern cells cannot themselves contain annotations.
            int linkCount = currentPage.Annotations.Count - annotationStart;
            if (linkCount == 0) return;
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeTransform> tiles = pattern.GetTileTransforms();
            drawingPatternAnnotations.Reserve(tiles.Count, linkCount);
            var links = currentPage.Annotations.GetRange(annotationStart, linkCount);
            currentPage.Annotations.RemoveRange(annotationStart, linkCount);
            foreach (OfficeTransform tile in tiles) {
                cancellationToken.ThrowIfCancellationRequested();
                int linkStart = currentPage.Annotations.Count;
                OfficeTransform transform = OfficeTransform.Scale(1D, -1D)
                    .Then(OfficeTransform.Translate(0D, pattern.InnerTile.Height))
                    .Then(tile).Then(new OfficeTransform(1D, 0D, 0D, -1D, originX, originTopY));
                foreach (LinkAnnotation link in links) {
                    cancellationToken.ThrowIfCancellationRequested();
                    currentPage.Annotations.Add(new LinkAnnotation {
                        X1 = link.X1, X2 = link.X2, Y1 = link.Y1, Y2 = link.Y2,
                        Uri = link.Uri, DestinationName = link.DestinationName, Contents = link.Contents,
                        StructElementIndex = link.StructElementIndex
                    });
                }
                TransformCanvasRectangles(currentPage.Annotations, linkStart, transform);
                ClipCanvasLinkAnnotations(currentPage.Annotations, linkStart, effect.BoundsLeft, effect.BoundsBottom,
                    pattern.Area.Width, pattern.Area.Height, OfficeClipPath.Rectangle(pattern.Area.Width, pattern.Area.Height));
            }
        }
    }
}
