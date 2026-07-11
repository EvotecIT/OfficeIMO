using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AddBackgroundDrawingPattern(
        ICollection<HtmlRenderVisual> visuals,
        OfficeDrawing drawing,
        OfficeImagePatternLayout pattern,
        int maximumTileCount,
        string visualSourceDescription) {
        var tiles = new List<HtmlRenderVisual>();
        foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(maximumTileCount)) {
            tiles.Add(HtmlRenderDrawing.CreateShared(
                drawing,
                tile.X,
                tile.Y,
                tile.Width,
                tile.Height,
                tiles.Count,
                alternativeText: null,
                linkUri: null,
                source: visualSourceDescription));
        }
        if (tiles.Count == 0) return;
        visuals.Add(new HtmlRenderClipGroup(
            pattern.Area.X,
            pattern.Area.Y,
            pattern.Area.Width,
            pattern.Area.Height,
            clipHorizontal: true,
            clipVertical: true,
            tiles,
            visuals.Count,
            visualSourceDescription + ":pattern-clip"));
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
