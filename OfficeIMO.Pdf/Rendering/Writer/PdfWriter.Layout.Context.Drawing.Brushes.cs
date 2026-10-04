using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private long _drawingTileCount;
        private int _drawingTileLimit;

        private void DrawDrawingPatternAt(OfficeDrawingTilingPattern pattern, double originX, double originTopY, OfficeDrawingTextMetrics metrics) {
            if (pattern.Opacity <= 0D) return;
            _drawingTileLimit = Math.Max(_drawingTileLimit, pattern.MaximumTileCount);
            RenderEffectGroup(OfficeTransform.Identity, pattern.Opacity, () => {
                var area = pattern.Area;
                new ContentStreamBuilder(sb).SaveState();
                AppendClipPath(sb, OfficeClipPath.Rectangle(area.Width, area.Height),
                    originX + area.X, originTopY - area.Y - area.Height, area.Height);
                foreach (OfficeTransform tileTransform in pattern.GetTileTransforms()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++_drawingTileCount > _drawingTileLimit)
                        throw new InvalidOperationException("PDF vector pattern aggregate expansion exceeds the configured tile-count limit.");
                    RenderEffectGroup(ToTopLeftPageTransform(tileTransform, originX, originTopY), 1D, () => {
                        new ContentStreamBuilder(sb).SaveState();
                        AppendClipPath(sb, OfficeClipPath.Rectangle(pattern.InnerTile.Width, pattern.InnerTile.Height),
                            originX, originTopY - pattern.InnerTile.Height, pattern.InnerTile.Height);
                        DrawDrawingElements(pattern.InnerTile, originX, originTopY, metrics);
                        new ContentStreamBuilder(sb).RestoreState();
                    });
                }
                new ContentStreamBuilder(sb).RestoreState();
            });
        }

        private void DrawDrawingEffectAt(OfficeDrawingEffectGroup effect, double originX, double originTopY, OfficeDrawingTextMetrics metrics) {
            OfficeTransform transform = ToTopLeftPageTransform(effect.Transform, originX, originTopY);
            PageEffectGroup? maskGroup = null;
            if (effect.SoftMask != null) {
                OfficeDrawingSoftMask mask = effect.SoftMask;
                if (mask.Mode != OfficeSoftMaskMode.Alpha || mask.BackdropColor.A != 0) {
                    throw new NotSupportedException("PDF drawing masks require alpha mode and a transparent backdrop.");
                }
                int start = sb.Length;
                bool previousWrappers = _suppressCanvasAccessibilityWrappers;
                bool previousText = _suppressCanvasActualTextChildren;
                OfficeTransform previousEffectToPage = _canvasEffectToPage;
                _suppressCanvasAccessibilityWrappers = true;
                _suppressCanvasActualTextChildren = true;
                _canvasEffectToPage = ConvertTopLeftCanvasTransform(transform, currentOpts.PageHeight).Then(previousEffectToPage);
                try {
                    RenderEffectGroup(OfficeTransform.Identity, 1D, OfficeBlendMode.Normal, () =>
                        RenderEffectGroup(ToTopLeftPageTransform(mask.Transform, originX, originTopY), 1D,
                            () => DrawDrawingElements(mask.InnerDrawing, originX, originTopY, metrics)), forceForm: true);
                    maskGroup = currentPage!.EffectGroups[currentPage.EffectGroups.Count - 1];
                } finally {
                    sb.Length = start;
                    _suppressCanvasAccessibilityWrappers = previousWrappers;
                    _suppressCanvasActualTextChildren = previousText;
                    _canvasEffectToPage = previousEffectToPage;
                }
            }
            int contentStart = sb.Length;
            RenderEffectGroup(transform, maskGroup == null ? effect.Opacity : 1D,
                maskGroup == null ? effect.BlendMode : OfficeBlendMode.Normal,
                () => DrawDrawingElements(effect.InnerDrawing, originX, originTopY, metrics), forceForm: maskGroup != null);
            if (maskGroup != null) {
                currentPage!.EffectGroups[currentPage.EffectGroups.Count - 1].AlphaMask = maskGroup;
                // Composite the completed masked result before applying group opacity.
                // Sharing its alpha state with the mask invocation attenuates the mask
                // as well as the result in PDF consumers.
                if (effect.Opacity < 1D || effect.BlendMode != OfficeBlendMode.Normal) {
                    string maskedContent = sb.ToString(contentStart, sb.Length - contentStart);
                    sb.Length = contentStart;
                    RenderEffectGroup(OfficeTransform.Identity, effect.Opacity, effect.BlendMode,
                        () => sb.Append(maskedContent), forceForm: true);
                }
            }
        }
    }
}
