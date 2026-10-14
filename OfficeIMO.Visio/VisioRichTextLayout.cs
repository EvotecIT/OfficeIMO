using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Shared layout and background placement for native character-run adapters.</summary>
internal static class VisioRichTextLayout {
    internal static OfficeRichTextBlockLayout Create(VisioRichTextProjection projection, double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, double minimumSize,
        CancellationToken cancellationToken) => projection.Paragraphs.Count > 0
        ? OfficeDrawingTextLayout.CreateParagraphs(projection.Paragraphs, width, height, measure,
            cancellationToken: cancellationToken, shrinkToFit: true, minimumFontSize: minimumSize)
        : OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
            projection.Runs, width, height, 1.2D, measure, wrap: true, shrinkToFit: true,
            minimumFontSize: minimumSize, cancellationToken: cancellationToken, shrinkToHeight: true);

    internal static OfficeTextBlockBackgroundBounds Background(OfficeRichTextBlockLayout layout,
        VisioRichTextProjection projection, VisioTextStyle? style, double x, double y, double width, double height,
        double paddingX, double paddingY) {
        double top = OfficeTextPlacement.ResolveTop(y - height / 2D, height, layout.Height,
            VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment));
        if (projection.Paragraphs.Count > 0)
            return OfficeDrawingTextLayout.CreateParagraphBackgroundBounds(layout,
                x - width / 2D, top, paddingX, paddingY);
        return new OfficeTextBlockBackgroundBounds(
            OfficeTextPlacement.ResolveLineLeft(x - width / 2D, width, layout.Width, projection.RenderAlignment) - paddingX,
            top - paddingY, layout.Width + 2D * paddingX, layout.Height + 2D * paddingY);
    }
}
