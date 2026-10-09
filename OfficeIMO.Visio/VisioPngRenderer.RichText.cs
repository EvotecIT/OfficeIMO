using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static partial class VisioPngRenderer {
    private static void DrawRichText(RasterCanvas canvas, VisioRichTextProjection projection,
        double x, double y, VisioTextStyle? style, double width, double height, double rotationRadians,
        bool drawLabelBackground) {
        OfficeRichTextBlockLayout layout = VisioRichTextLayout.Create(projection, width, height,
            canvas.MeasureText, canvas.Supersampling * 5D, canvas.CancellationToken);
        if (ResolveTextBackground(style, drawLabelBackground) is OfficeColor background && background.A > 0) {
            OfficeTextBlockBackgroundBounds bounds = VisioRichTextLayout.Background(layout, projection, style,
                x, y, width, height, canvas.Supersampling * 3D, canvas.Supersampling * 2D);
            canvas.DrawRichTextBackground(bounds, background, rotationRadians, x, y);
        }
        canvas.DrawRichText(layout, x - width / 2D, y - height / 2D, width, height,
            projection.RenderAlignment, VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment),
            rotationRadians, x, y);
    }
}
