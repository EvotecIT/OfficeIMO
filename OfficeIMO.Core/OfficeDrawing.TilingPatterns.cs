using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Adds a clipped, transform-aware vector tiling pattern.</summary>
    public OfficeDrawing AddTilingPattern(
        OfficeDrawing tile,
        OfficeImagePlacement area,
        double horizontalStep,
        double verticalStep,
        OfficeTransform? transform = null,
        double originX = 0D,
        double originY = 0D,
        int maximumTileCount = 16384,
        double opacity = 1D) =>
        AddTilingPattern(tile, area, horizontalStep, verticalStep, true, true, transform, originX, originY, maximumTileCount, opacity);

    /// <summary>Adds a clipped, transform-aware vector tiling pattern with independent axis repetition.</summary>
    public OfficeDrawing AddTilingPattern(
        OfficeDrawing tile,
        OfficeImagePlacement area,
        double horizontalStep,
        double verticalStep,
        bool repeatX,
        bool repeatY,
        OfficeTransform? transform = null,
        double originX = 0D,
        double originY = 0D,
        int maximumTileCount = 16384,
        double opacity = 1D) {
        if (area.X < 0D || area.Y < 0D || area.X + area.Width > Width || area.Y + area.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(area), "Pattern area must fit inside the drawing bounds.");
        }
        if (tile == null) throw new ArgumentNullException(nameof(tile));
        Fonts.AddRange(tile.Fonts);
        _elements.Add(new OfficeDrawingTilingPattern(tile, area, horizontalStep, verticalStep, repeatX, repeatY, transform, originX, originY, maximumTileCount, opacity));
        return this;
    }

    private void AddNestedTilingPattern(OfficeDrawingTilingPattern pattern, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        OfficeImagePlacement area = pattern.Area;
        var translatedArea = new OfficeImagePlacement(area.X + offsetX, area.Y + offsetY, area.Width, area.Height);
        if (!allowOverflow && (translatedArea.X < 0D || translatedArea.Y < 0D || translatedArea.X + translatedArea.Width > Width || translatedArea.Y + translatedArea.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(pattern), "Nested pattern area must fit inside the drawing bounds.");
        }
        OfficeTransform transform = pattern.Transform.Then(OfficeTransform.Translate(offsetX, offsetY));
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) transform = transform.Then(frameTransform.Value.CreateDestinationTransform());
        Fonts.AddRange(pattern.InnerTile.Fonts);
        _elements.Add(new OfficeDrawingTilingPattern(
            pattern.InnerTile,
            translatedArea,
            pattern.HorizontalStep,
            pattern.VerticalStep,
            pattern.RepeatX,
            pattern.RepeatY,
            transform,
            pattern.OriginX,
            pattern.OriginY,
            pattern.MaximumTileCount,
            pattern.Opacity));
    }
}
