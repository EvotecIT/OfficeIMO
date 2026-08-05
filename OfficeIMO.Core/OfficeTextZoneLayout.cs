using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes reusable bounded text zones inside a horizontal rendering area.
/// </summary>
public sealed class OfficeTextZoneLayout {
    private OfficeTextZoneLayout(OfficeTextZone left, OfficeTextZone center, OfficeTextZone right) {
        Left = left;
        Center = center;
        Right = right;
    }

    /// <summary>
    /// Gets the left-aligned text zone.
    /// </summary>
    public OfficeTextZone Left { get; }

    /// <summary>
    /// Gets the centered text zone.
    /// </summary>
    public OfficeTextZone Center { get; }

    /// <summary>
    /// Gets the right-aligned text zone.
    /// </summary>
    public OfficeTextZone Right { get; }

    /// <summary>
    /// Creates three non-overlapping horizontal text zones with bounded padding and gaps.
    /// </summary>
    /// <param name="width">Total available width.</param>
    /// <param name="padding">Requested horizontal padding.</param>
    /// <param name="gap">Requested gap between adjacent zones.</param>
    /// <returns>A layout containing left, center, and right text zones.</returns>
    public static OfficeTextZoneLayout CreateThreeColumn(double width, double padding, double gap) {
        double boundedWidth = NormalizePositive(width, 1D);
        double boundedPadding = Math.Max(0D, Math.Min(NormalizeNonNegative(padding), boundedWidth / 6D));
        double boundedGap = Math.Max(0D, Math.Min(NormalizeNonNegative(gap), boundedWidth / 24D));
        double contentLeft = boundedPadding;
        double contentRight = Math.Max(contentLeft, boundedWidth - boundedPadding);
        double contentWidth = Math.Max(1D, contentRight - contentLeft);
        double zoneWidth = Math.Max(1D, (contentWidth - (boundedGap * 2D)) / 3D);
        double leftX = contentLeft;
        double centerX = leftX + zoneWidth + boundedGap;
        double rightX = centerX + zoneWidth + boundedGap;
        return new OfficeTextZoneLayout(
            new OfficeTextZone(leftX, zoneWidth, OfficeTextPlacement.ResolveAnchorX(leftX, zoneWidth, OfficeTextAlignment.Left)),
            new OfficeTextZone(centerX, zoneWidth, OfficeTextPlacement.ResolveAnchorX(centerX, zoneWidth, OfficeTextAlignment.Center)),
            new OfficeTextZone(rightX, zoneWidth, OfficeTextPlacement.ResolveAnchorX(rightX, zoneWidth, OfficeTextAlignment.Right)));
    }

    private static double NormalizePositive(double value, double fallback) =>
        value > 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0D;
}

/// <summary>
/// Describes one bounded text zone and the anchor coordinate for its text alignment.
/// </summary>
public readonly struct OfficeTextZone {
    /// <summary>
    /// Initializes a new bounded text zone.
    /// </summary>
    /// <param name="x">Left edge of the zone.</param>
    /// <param name="width">Zone width.</param>
    /// <param name="anchorX">Text anchor X coordinate.</param>
    public OfficeTextZone(double x, double width, double anchorX) {
        X = x;
        Width = Math.Max(0D, width);
        AnchorX = anchorX;
    }

    /// <summary>
    /// Gets the left edge of the zone.
    /// </summary>
    public double X { get; }

    /// <summary>
    /// Gets the zone width.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Gets the text anchor X coordinate for the zone.
    /// </summary>
    public double AnchorX { get; }
}
