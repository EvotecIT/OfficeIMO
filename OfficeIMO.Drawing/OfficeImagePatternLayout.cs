using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes a clipped, regularly tiled image pattern in drawing coordinates.
/// </summary>
public readonly struct OfficeImagePatternLayout {
    /// <summary>Creates a pattern layout from a paint area and one positioned tile.</summary>
    public OfficeImagePatternLayout(OfficeImagePlacement area, OfficeImagePlacement tile, bool repeatX = true, bool repeatY = true) {
        EnsureFinite(area.X + area.Width, nameof(area));
        EnsureFinite(area.Y + area.Height, nameof(area));
        EnsureFinite(tile.X + tile.Width, nameof(tile));
        EnsureFinite(tile.Y + tile.Height, nameof(tile));
        Area = area;
        Tile = tile;
        RepeatX = repeatX;
        RepeatY = repeatY;
    }

    /// <summary>Clipped destination area painted by the pattern.</summary>
    public OfficeImagePlacement Area { get; }

    /// <summary>Position and size of the authored origin tile.</summary>
    public OfficeImagePlacement Tile { get; }

    /// <summary>Whether the origin tile repeats horizontally.</summary>
    public bool RepeatX { get; }

    /// <summary>Whether the origin tile repeats vertically.</summary>
    public bool RepeatY { get; }

    /// <summary>Number of tiles intersecting the clipped area, saturated at <see cref="long.MaxValue" />.</summary>
    public long EstimatedTileCount {
        get {
            AxisPlan horizontal = ResolveAxis(Area.X, Area.Width, Tile.X, Tile.Width, RepeatX);
            AxisPlan vertical = ResolveAxis(Area.Y, Area.Height, Tile.Y, Tile.Height, RepeatY);
            if (horizontal.Count == 0L || vertical.Count == 0L) return 0L;
            return horizontal.Count > long.MaxValue / vertical.Count
                ? long.MaxValue
                : horizontal.Count * vertical.Count;
        }
    }

    /// <summary>Returns all full tile placements intersecting the clipped area within an explicit safety limit.</summary>
    public IReadOnlyList<OfficeImagePlacement> GetTilePlacements(int maximumTileCount) {
        if (maximumTileCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumTileCount), "Maximum image-pattern tile count must be positive.");
        }

        AxisPlan horizontal = ResolveAxis(Area.X, Area.Width, Tile.X, Tile.Width, RepeatX);
        AxisPlan vertical = ResolveAxis(Area.Y, Area.Height, Tile.Y, Tile.Height, RepeatY);
        long count = horizontal.Count == 0L || vertical.Count == 0L
            ? 0L
            : horizontal.Count > long.MaxValue / vertical.Count
                ? long.MaxValue
                : horizontal.Count * vertical.Count;
        if (count > maximumTileCount) {
            throw new InvalidOperationException("Image pattern exceeds the configured tile-count limit.");
        }

        var placements = new List<OfficeImagePlacement>((int)count);
        for (long row = 0L; row < vertical.Count; row++) {
            double y = vertical.First + (row * Tile.Height);
            for (long column = 0L; column < horizontal.Count; column++) {
                placements.Add(new OfficeImagePlacement(
                    horizontal.First + (column * Tile.Width),
                    y,
                    Tile.Width,
                    Tile.Height));
            }
        }

        return new ReadOnlyCollection<OfficeImagePlacement>(placements);
    }

    /// <summary>Returns a translated pattern layout.</summary>
    public OfficeImagePatternLayout Translate(double offsetX, double offsetY) {
        EnsureFinite(offsetX, nameof(offsetX));
        EnsureFinite(offsetY, nameof(offsetY));
        return new OfficeImagePatternLayout(
            new OfficeImagePlacement(Area.X + offsetX, Area.Y + offsetY, Area.Width, Area.Height),
            new OfficeImagePlacement(Tile.X + offsetX, Tile.Y + offsetY, Tile.Width, Tile.Height),
            RepeatX,
            RepeatY);
    }

    /// <summary>Returns a uniformly scaled pattern layout.</summary>
    public OfficeImagePatternLayout Scale(double scale) {
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Image-pattern scale must be finite and positive.");
        }

        return new OfficeImagePatternLayout(
            new OfficeImagePlacement(Area.X * scale, Area.Y * scale, Area.Width * scale, Area.Height * scale),
            new OfficeImagePlacement(Tile.X * scale, Tile.Y * scale, Tile.Width * scale, Tile.Height * scale),
            RepeatX,
            RepeatY);
    }

    private static AxisPlan ResolveAxis(double areaStart, double areaLength, double tileStart, double tileLength, bool repeat) {
        double areaEnd = areaStart + areaLength;
        if (!repeat) {
            return tileStart < areaEnd && tileStart + tileLength > areaStart
                ? new AxisPlan(tileStart, 1L)
                : new AxisPlan(tileStart, 0L);
        }

        double first = tileStart + (Math.Floor((areaStart - tileStart) / tileLength) * tileLength);
        if (first + tileLength <= areaStart + 0.0000001D) first += tileLength;
        if (first >= areaEnd - 0.0000001D) return new AxisPlan(first, 0L);
        double required = Math.Ceiling((areaEnd - first) / tileLength);
        long count = required >= long.MaxValue ? long.MaxValue : Math.Max(0L, (long)required);
        return new AxisPlan(first, count);
    }

    private static void EnsureFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Image-pattern offsets must be finite.");
        }
    }

    private readonly struct AxisPlan {
        internal AxisPlan(double first, long count) {
            First = first;
            Count = count;
        }

        internal double First { get; }
        internal long Count { get; }
    }
}
