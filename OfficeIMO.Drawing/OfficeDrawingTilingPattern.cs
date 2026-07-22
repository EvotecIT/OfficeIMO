using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Clipped, transform-aware repetition of a reusable vector drawing tile.
/// </summary>
public sealed class OfficeDrawingTilingPattern : OfficeDrawingElement {
    private readonly OfficeDrawing _tile;

    /// <summary>Creates a bounded vector tiling pattern.</summary>
    public OfficeDrawingTilingPattern(
        OfficeDrawing tile,
        OfficeImagePlacement area,
        double horizontalStep,
        double verticalStep,
        OfficeTransform? transform = null,
        double originX = 0D,
        double originY = 0D,
        int maximumTileCount = 16384,
        double opacity = 1D) {
        if (tile == null) throw new ArgumentNullException(nameof(tile));
        if (tile.Width <= 0D || tile.Height <= 0D) throw new ArgumentException("Pattern tile dimensions must be positive.", nameof(tile));
        if (area.Width <= 0D || area.Height <= 0D) throw new ArgumentOutOfRangeException(nameof(area), "Pattern area dimensions must be positive.");
        HorizontalStep = ValidateStep(horizontalStep, nameof(horizontalStep));
        VerticalStep = ValidateStep(verticalStep, nameof(verticalStep));
        EnsureFinite(originX, nameof(originX));
        EnsureFinite(originY, nameof(originY));
        if (maximumTileCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTileCount));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) throw new ArgumentOutOfRangeException(nameof(opacity));

        Transform = transform ?? OfficeTransform.Identity;
        if (!Transform.TryInvert(out _)) throw new ArgumentException("Pattern transform must be invertible.", nameof(transform));
        _tile = tile.Clone();
        Area = area;
        OriginX = originX;
        OriginY = originY;
        MaximumTileCount = maximumTileCount;
        Opacity = opacity;
        _ = GetTileTransforms(maximumTileCount);
    }

    /// <summary>Detached vector tile drawing.</summary>
    public OfficeDrawing Tile => _tile.Clone();

    /// <summary>Destination-space clipping area.</summary>
    public OfficeImagePlacement Area { get; }

    /// <summary>Horizontal distance between tile origins in pattern space.</summary>
    public double HorizontalStep { get; }

    /// <summary>Vertical distance between tile origins in pattern space.</summary>
    public double VerticalStep { get; }

    /// <summary>Pattern-space affine transform.</summary>
    public OfficeTransform Transform { get; }

    /// <summary>Pattern-space horizontal grid origin.</summary>
    public double OriginX { get; }

    /// <summary>Pattern-space vertical grid origin.</summary>
    public double OriginY { get; }

    /// <summary>Maximum number of tiles expanded by renderers.</summary>
    public int MaximumTileCount { get; }

    /// <summary>Pattern opacity from zero through one.</summary>
    public double Opacity { get; }

    internal OfficeDrawing InnerTile => _tile;

    /// <summary>Returns the bounded transforms needed to paint this pattern.</summary>
    public IReadOnlyList<OfficeTransform> GetTileTransforms() => GetTileTransforms(MaximumTileCount);

    internal IReadOnlyList<OfficeTransform> GetTileTransforms(int maximumTileCount) {
        if (maximumTileCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTileCount));
        Transform.TryInvert(out OfficeTransform inverse);
        OfficePoint topLeft = inverse.TransformPoint(new OfficePoint(Area.X, Area.Y));
        OfficePoint topRight = inverse.TransformPoint(new OfficePoint(Area.X + Area.Width, Area.Y));
        OfficePoint bottomRight = inverse.TransformPoint(new OfficePoint(Area.X + Area.Width, Area.Y + Area.Height));
        OfficePoint bottomLeft = inverse.TransformPoint(new OfficePoint(Area.X, Area.Y + Area.Height));
        double minX = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        double maxX = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        double minY = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        double maxY = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
        long firstColumn = (long)Math.Floor((minX - OriginX - _tile.Width) / HorizontalStep) + 1L;
        long lastColumn = (long)Math.Ceiling((maxX - OriginX) / HorizontalStep) - 1L;
        long firstRow = (long)Math.Floor((minY - OriginY - _tile.Height) / VerticalStep) + 1L;
        long lastRow = (long)Math.Ceiling((maxY - OriginY) / VerticalStep) - 1L;
        long columns = Math.Max(0L, lastColumn - firstColumn + 1L);
        long rows = Math.Max(0L, lastRow - firstRow + 1L);
        long count = columns == 0L || rows == 0L || columns > long.MaxValue / rows ? (columns == 0L || rows == 0L ? 0L : long.MaxValue) : columns * rows;
        if (count > maximumTileCount) throw new InvalidOperationException("Vector pattern exceeds the configured tile-count limit.");

        var result = new List<OfficeTransform>((int)count);
        for (long row = firstRow; row <= lastRow; row++) {
            for (long column = firstColumn; column <= lastColumn; column++) {
                result.Add(OfficeTransform.Translate(
                    OriginX + (column * HorizontalStep),
                    OriginY + (row * VerticalStep)).Then(Transform));
            }
        }
        return new ReadOnlyCollection<OfficeTransform>(result);
    }

    internal override OfficeDrawingElement CloneElement() => new OfficeDrawingTilingPattern(
        _tile, Area, HorizontalStep, VerticalStep, Transform, OriginX, OriginY, MaximumTileCount, Opacity);

    private static double ValidateStep(double value, string parameterName) {
        EnsureFinite(value, parameterName);
        if (value == 0D) throw new ArgumentOutOfRangeException(parameterName, "Pattern steps must be non-zero.");
        return Math.Abs(value);
    }

    private static void EnsureFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(parameterName);
    }
}
