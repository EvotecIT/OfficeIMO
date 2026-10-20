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
                rotationDegrees: -item.RotationAngle,
                rotationCenterX: item.X + block.Width / 2D,
                rotationCenterY: item.Y + block.Height / 2D,
                flipHorizontal: item.HorizontalFlip, flipVertical: item.VerticalFlip);

            if (item.Foreground) {
                // The foreground image queue owns image streams, not vector forms.
                // Preserve its z-order and transform using the same bounded SVG fallback.
                byte[] png = RasterizeSvgImageForPdf(ReadSvgImageForPdf(block.Data), projection);
                currentOpts.AddLayoutDiagnostic("SvgImageRasterFallback", "Canvas image",
                    "The SVG image was rasterized for the foreground image layer.", PdfLayoutDiagnosticKind.SimplifiedContent);
                PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(png, cancellationToken);
                var raster = new ImageBlock(prepared.Data, block.Width, block.Height, prepared.Info, style,
                    block.LinkUri, block.LinkContents, useDataSnapshot: true, preparedStream: prepared.PreparedStream);
                RenderCanvasImage(new PdfCanvasImageItem(raster, item.X, item.Y, item.RotationAngle,
                    item.HorizontalFlip, item.VerticalFlip, foreground: true, foregroundZOrder: item.ForegroundZOrder, logicalOrder: item.LogicalOrder));
                return;
            }

            OfficeClipPath? clip = style.ClipPath;
            if (clip == null && plan.RequiresTargetClip) clip = OfficeClipPath.Rectangle(block.Width, block.Height);
            if (clip != null) {
                new ContentStreamBuilder(sb).SaveState();
                AppendClipPath(sb, clip, item.X, bottomY, block.Height);
            }
            try {
                OfficeDrawingImage image = OfficeDrawingImage.CreateWithInterpolation(block.Data, "image/svg+xml",
                    projection, interpolate: true, alternativeText: style.AlternativeText);
                TryDrawSvgImageAt(image, 0D, currentOpts.PageHeight, item.LogicalOrder);
            } finally {
                if (clip != null) new ContentStreamBuilder(sb).RestoreState();
            }

            PageImage geometry = CreatePageImage(block, style, item.X, bottomY, block.Width, block.Height);
            geometry.RotationAngle = item.RotationAngle;
            geometry.HorizontalFlip = item.HorizontalFlip;
            geometry.VerticalFlip = item.VerticalFlip;
            AddImageLinkAnnotation(block, style, geometry, item.X, bottomY, block.Width, block.Height);
            if (paragraphCanvasTop.HasValue) ReportAnchoredImagePageClipping(item, bottomY);
            DrawDebugCanvasItemBox(item.X, bottomY, block.Width, block.Height);
            pageDirty = true;
        }
    }
}
