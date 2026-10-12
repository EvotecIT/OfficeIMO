using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private (double Left, double Top, double Right, double Bottom) GroupChildPaintBounds(OfficeDrawing local) {
        double left = 0, top = 0, right = local.Width, bottom = local.Height;
        // Pictures retain their nominal frame. Shapes can paint beyond it:
        // custom paths, stroke joins, line markers and shadows all need extra room.
        foreach (OfficeDrawingShape item in local.Elements.OfType<OfficeDrawingShape>()) {
            _context.Token.ThrowIfCancellationRequested();
            OfficeShape shape = item.Shape;
            // An unspecified join can use PDF's implicit miter limit of ten.
            double miterLimit = shape.StrokeLineJoin.HasValue ? shape.StrokeMiterLimit : Math.Max(10, shape.StrokeMiterLimit);
            bool paintsStroke = shape.StrokeOpacity.GetValueOrDefault(1) > 0 && (shape.StrokeGradient != null || shape.StrokeRadialGradient != null
                || shape.StrokeColor is OfficeColor { A: > 0 });
            // Disabled native lines retain their stored width but paint no stroke.
            double padding = 1 + (paintsStroke ? Math.Max(0, shape.StrokeWidth) / 2 * Math.Max(2, miterLimit) : 0);
            foreach (OfficeShadowLayer layer in OfficeDrawingRasterRenderer.CreateShadowLayers(shape)) {
                OfficeShadow shadow = shape.Shadow!;
                // Use the renderer's actual expansion/stroke plan, retaining
                // square caps and soft miter joins without reserving hidden ink.
                double shadowPadding = layer.Expansion + (layer.HasStroke ? layer.StrokeWidth / 2 * Math.Max(2, miterLimit) : 0);
                padding = Math.Max(padding, 1 + shadowPadding + Math.Max(Math.Abs(shadow.OffsetX), Math.Abs(shadow.OffsetY)));
            }
            if (paintsStroke)
                foreach (OfficeLineMarker? marker in new[] { shape.StrokeStartMarker, shape.StrokeEndMarker })
                    if (marker != null) padding += Math.Max(marker.Width, marker.Length);
            OfficeTransform placement = (shape.Transform ?? OfficeTransform.Identity)
                .Then(OfficeTransform.Translate(item.X, item.Y));
            foreach (OfficeFlattenedPathContour contour in OfficeStrokeGeometry.FlattenShape(shape, 4))
                foreach (OfficePoint point in contour.Points) {
                    var bounds = placement.TransformRectangleBounds(point.X - padding, point.Y - padding,
                        padding * 2, padding * 2);
                    left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                    right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                }
        }
        // Text layout bounds include glyph overhang, decorations and condensed
        // line ink. Reuse the raster owner's metrics without changing text flow
        // or expanding explicit clips on nested picture/table drawings.
        foreach (OfficeDrawingRichText text in local.Elements.OfType<OfficeDrawingRichText>()) {
            double width = text.Width - text.Padding.Horizontal, height = text.Height - text.Padding.Vertical;
            if (width <= 0 || height <= 0) continue;
            _context.AccountTextLayout(text.Paragraphs.Sum(paragraph => paragraph.Runs.Sum(run => run.Text.Length)));
            _textMetrics ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, null,
                cancellationToken: _context.Token);
            OfficeRasterCanvas metrics = _textMetrics.WithDrawingTextProfile(local);
            OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, width, height, metrics);
            foreach (var ink in OfficeDrawingTextLayout.PlacedRasterParagraphPaintBounds(layout, metrics,
                text.X + text.Padding.Left, text.Y + text.Padding.Top, height, text.VerticalAlignment)) {
                var bounds = text.HasFrameTransform
                    ? text.CreateFrameTransform().CreateDestinationTransform().TransformRectangleBounds(
                        ink.Left, ink.Top, ink.Right - ink.Left, ink.Bottom - ink.Top) : ink;
                left = Math.Min(left, bounds.Left - 1); top = Math.Min(top, bounds.Top - 1);
                right = Math.Max(right, bounds.Right + 1); bottom = Math.Max(bottom, bounds.Bottom + 1);
            }
        }
        // Keep an integral local raster origin so added overhang room does not
        // shift the original text sampling grid.
        return (Math.Floor(left), Math.Floor(top), Math.Ceiling(right), Math.Ceiling(bottom));
    }
}
