using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Resolve all sized text frames before paint. Connectors may precede their
    // targets or live in another group; the attachment owner needs the complete
    // projection bounds, not an order-dependent partially populated overlay.
    private static void PrepareTextFrames(OdgShapes shapes, OfficeDrawing drawing, OdfConversionReport report,
        OdgLayers layers, bool forPrint, CancellationToken cancellationToken,
        DrawingFieldContext fields, OfficeDrawingTextMetrics? layoutMetrics) {
        foreach (OdgShape shape in shapes) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (shape.Layer is string layerName && layers.Find(layerName) is OdgLayer layer && !layer.IsVisible(forPrint)) continue;
                // Invalid group transforms omit that subtree in the paint pass.
                OdfDrawingTransform.Parse(shape.Transform);
                if (shape.IsGroup) {
                    PrepareTextFrames(shape.Children, drawing, report, layers, forPrint, cancellationToken, fields, layoutMetrics);
                    continue;
                }
                if (!PrepareTextBoxLayout(shape)) continue;
                OdfRect bounds = shape.Bounds;
                double width = bounds.Width.ToPoints(), height = bounds.Height.ToPoints();
                var local = new OfficeDrawing(Math.Max(width, .001D), Math.Max(height, .001D));
                CopyDrawingResources(drawing, local);
                TextFrameProjection prepared = ProjectText(shape, local, report, width, height, fields,
                    cancellationToken: cancellationToken, layoutMetrics: layoutMetrics);
                fields.PreparedTextFrames.Add(shape.Element, prepared);
                fields.ProjectedFrameBounds.Add(shape.Element, new OdfRect(bounds.X, bounds.Y,
                    OdfLength.Points(prepared.Width), OdfLength.Points(prepared.Height)));
            } catch (Exception exception) when (exception is FormatException or ArgumentException or NotSupportedException or InvalidDataException or OverflowException) {
                // The paint pass owns the geometry failure mapping. Cancellation
                // and unexpected failures propagate without a partial result.
            }
        }
    }
}
