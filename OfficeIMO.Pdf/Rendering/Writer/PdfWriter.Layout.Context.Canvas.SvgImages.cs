using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Projects canvas SVG images through the existing Drawing SVG writer.</summary>
        private void RenderCanvasSvgImage(PdfCanvasImageItem item, PdfImageStyle style, double bottomY) {
            ImageBlock block = item.Block;
            OfficeImageRenderPlan plan = OfficeImageRenderPlan.CreateTopLeft(block.Info.Width, block.Info.Height,
                item.X, item.Y, block.Width, block.Height, style.Fit, style.SourceCrop?.ToOfficeImageSourceCrop() ?? default);
            OfficeImageProjection projection = plan.ToVisibleProjection(
                rotationDegrees: -item.RotationAngle - style.RotationAngle,
                rotationCenterX: item.X + block.Width / 2D,
                rotationCenterY: item.Y + block.Height / 2D,
                flipHorizontal: item.HorizontalFlip ^ style.HorizontalFlip, flipVertical: item.VerticalFlip ^ style.VerticalFlip);

            OfficeClipPath? clip = style.ClipPath;
            if (clip == null && plan.RequiresTargetClip) clip = OfficeClipPath.Rectangle(block.Width, block.Height);
            if (clip != null) {
                new ContentStreamBuilder(sb).SaveState();
                AppendClipPath(sb, clip, item.X, bottomY, block.Height);
            }
            try {
                // The block owns an immutable snapshot and, from admission, its parsed drawing.
                var image = new OfficeDrawingImage(block.Data, "image/svg+xml", projection, style.AltText, opacity: 1D,
                    interpolate: true, useDataSnapshot: true);
                TryDrawSvgImageAt(image, 0D, currentOpts.PageHeight.ToPoints(), item.LogicalOrder, block.SvgDrawing);
            } finally {
                if (clip != null) new ContentStreamBuilder(sb).RestoreState();
            }

            PageImage geometry = CreatePageImage(block, style, item.X, bottomY, block.Width, block.Height);
            geometry.RotationAngle += item.RotationAngle;
            geometry.HorizontalFlip ^= item.HorizontalFlip;
            geometry.VerticalFlip ^= item.VerticalFlip;
            AddImageLinkAnnotation(block, style, geometry, item.X, bottomY, block.Width, block.Height);
            if (paragraphCanvasTop.HasValue) ReportAnchoredImagePageClipping(item, bottomY);
            DrawDebugCanvasItemBox(item.X, bottomY, block.Width, block.Height);
            pageDirty = true;
        }
    }
}
