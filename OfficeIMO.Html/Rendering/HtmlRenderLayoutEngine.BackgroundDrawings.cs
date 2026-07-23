using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AddBackgroundDrawingPattern(
        ICollection<HtmlRenderVisual> visuals,
        OfficeDrawing drawing,
        OfficeImagePatternLayout pattern,
        int maximumTileCount,
        string visualSourceDescription) {
        if (pattern.EstimatedTileCount <= 0L || pattern.EstimatedTileCount > maximumTileCount) return;
        OfficeImagePlacement tile = pattern.Tile;
        var tileDrawing = new OfficeDrawing(tile.Width, tile.Height);
        tileDrawing.AddEffectDrawing(drawing, OfficeTransform.Scale(
            tile.Width / drawing.Width,
            tile.Height / drawing.Height));

        OfficeImagePlacement area = pattern.Area;
        var patternDrawing = new OfficeDrawing(area.Width, area.Height);
        patternDrawing.AddTilingPattern(
            tileDrawing,
            new OfficeImagePlacement(0D, 0D, area.Width, area.Height),
            pattern.HorizontalStep,
            pattern.VerticalStep,
            originX: tile.X - area.X,
            originY: tile.Y - area.Y,
            maximumTileCount: maximumTileCount,
            repeatX: pattern.RepeatX,
            repeatY: pattern.RepeatY);
        visuals.Add(HtmlRenderDrawing.CreateShared(
            patternDrawing,
            area.X,
            area.Y,
            area.Width,
            area.Height,
            visuals.Count,
            alternativeText: null,
            linkUri: null,
            source: visualSourceDescription));
    }

    private static void AddVisibleBackgroundDrawing(
        ICollection<HtmlRenderVisual> visuals,
        OfficeDrawing drawing,
        double tileX,
        double tileY,
        double tileWidth,
        double tileHeight,
        double areaX,
        double areaY,
        double areaWidth,
        double areaHeight,
        string visualSourceDescription) {
        double visibleLeft = Math.Max(tileX, areaX);
        double visibleTop = Math.Max(tileY, areaY);
        double visibleRight = Math.Min(tileX + tileWidth, areaX + areaWidth);
        double visibleBottom = Math.Min(tileY + tileHeight, areaY + areaHeight);
        if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) return;

        HtmlRenderDrawing tile = HtmlRenderDrawing.CreateShared(
            drawing,
            tileX,
            tileY,
            tileWidth,
            tileHeight,
            0,
            alternativeText: null,
            linkUri: null,
            source: visualSourceDescription);
        visuals.Add(new HtmlRenderClipGroup(
            visibleLeft,
            visibleTop,
            visibleRight - visibleLeft,
            visibleBottom - visibleTop,
            clipHorizontal: true,
            clipVertical: true,
            new[] { tile },
            visuals.Count,
            visualSourceDescription + ":clip"));
    }
}
