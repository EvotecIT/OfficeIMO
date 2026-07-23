using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws a filled and/or stroked ellipse using a shared Office stroke dash style.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="fill">Fill color.</param>
    /// <param name="stroke">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashStyle">Shared Office stroke dash style.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate dashed outlines.</param>
    public void DrawStyledEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor fill,
        OfficeColor stroke,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72) {
        if (dashStyle == OfficeStrokeDashStyle.Solid || stroke.A == 0 || thickness <= 0D) {
            DrawEllipse(centerX, centerY, radiusX, radiusY, fill, stroke, thickness, rotationDegrees, rotationCenterX, rotationCenterY);
            return;
        }

        if (fill.A > 0) {
            DrawEllipse(centerX, centerY, radiusX, radiusY, fill, OfficeColor.Transparent, 0D, rotationDegrees, rotationCenterX, rotationCenterY);
        }

        DrawPatternedEllipse(
            centerX,
            centerY,
            radiusX,
            radiusY,
            stroke,
            thickness,
            dashStyle.GetDashPattern(thickness),
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            segments);
    }

    /// <summary>
    /// Draws a dashed elliptical outline using center/radius coordinates and optional rotation.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashLength">Visible dash length in canvas pixels.</param>
    /// <param name="gapLength">Transparent gap length in canvas pixels.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate the ellipse.</param>
    /// <param name="resetDashPatternForEachSegment">Whether the dash pattern should restart for every approximation segment.</param>
    public void DrawDashedEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor color,
        double thickness = 1D,
        double dashLength = 6D,
        double gapLength = 4D,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D || dashLength <= 0D || gapLength < 0D || segments < 4) {
            return;
        }

        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        double patternPosition = 0D;
        OfficePoint previous = CreateArcStartPoint(centerX, centerY, radiusX, radiusY, 0D, rotationRadians, rotationCenterX, rotationCenterY);
        foreach (OfficePoint current in OfficeGeometry.CreateEllipticalArcPoints(
            centerX,
            centerY,
            radiusX,
            radiusY,
            0D,
            Math.PI * 2D,
            segments,
            rotationRadians,
            rotationCenterX,
            rotationCenterY)) {
            if (resetDashPatternForEachSegment) {
                DrawDashedLine(previous.X, previous.Y, current.X, current.Y, color, thickness, dashLength, gapLength);
            } else {
                DrawDashedPathSegment(previous, current, color, thickness, dashLength, gapLength, ref patternPosition);
            }

            previous = current;
        }
    }

    /// <summary>
    /// Draws an elliptical outline using an alternating dash and gap pattern.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashPattern">Alternating dash and gap lengths in canvas pixels.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate the ellipse.</param>
    public void DrawPatternedEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double>? dashPattern,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72) {
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D || segments < 4) {
            return;
        }

        List<double> pattern = NormalizeDashPattern(dashPattern);
        if (pattern.Count == 0) {
            DrawEllipse(centerX, centerY, radiusX, radiusY, OfficeColor.Transparent, color, thickness, rotationDegrees, rotationCenterX, rotationCenterY);
            return;
        }

        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        double patternPosition = 0D;
        OfficePoint previous = CreateArcStartPoint(centerX, centerY, radiusX, radiusY, 0D, rotationRadians, rotationCenterX, rotationCenterY);
        foreach (OfficePoint current in OfficeGeometry.CreateEllipticalArcPoints(
            centerX,
            centerY,
            radiusX,
            radiusY,
            0D,
            Math.PI * 2D,
            segments,
            rotationRadians,
            rotationCenterX,
            rotationCenterY)) {
            DrawPatternedPathSegment(previous, current, color, thickness, pattern, ref patternPosition);
            previous = current;
        }
    }

    private void DrawDashedPathSegment(
        OfficePoint start,
        OfficePoint end,
        OfficeColor color,
        double thickness,
        double dashLength,
        double gapLength,
        ref double patternPosition) {
        NormalizeRasterDashLengths(ref dashLength, ref gapLength);
        double length = Distance(start.X, start.Y, end.X, end.Y);
        if (!IsFinite(length) || length <= 0D) {
            return;
        }

        double cycle = SaturatingDashCycle(dashLength, gapLength);

        OfficePoint clippedStart = start;
        OfficePoint clippedEnd = end;
        if (!TryClipLineToCanvas(ref clippedStart, ref clippedEnd, thickness, length, out double leadingDistance, out double trailingDistance)) {
            patternPosition = AdvancePatternPosition(patternPosition, length, cycle);
            return;
        }
        patternPosition = AdvancePatternPosition(patternPosition, leadingDistance, cycle);
        length = Distance(clippedStart.X, clippedStart.Y, clippedEnd.X, clippedEnd.Y);
        double position = 0D;
        while (position < length) {
            bool inDash = patternPosition < dashLength || gapLength == 0D;
            double patternRemaining = inDash
                ? dashLength - patternPosition
                : cycle - patternPosition;
            if (patternRemaining <= MinimumDashSegmentAdvance) {
                patternPosition = inDash ? dashLength : 0D;
                continue;
            }

            double next = Math.Min(length, position + patternRemaining);
            double consumed = next - position;
            if (consumed <= MinimumDashSegmentAdvance) {
                break;
            }

            if (inDash) {
                double startT = position / length;
                double endT = next / length;
                DrawLineSegment(
                    clippedStart.X + ((clippedEnd.X - clippedStart.X) * startT),
                    clippedStart.Y + ((clippedEnd.Y - clippedStart.Y) * startT),
                    clippedStart.X + ((clippedEnd.X - clippedStart.X) * endT),
                    clippedStart.Y + ((clippedEnd.Y - clippedStart.Y) * endT),
                    color,
                    thickness);
            }

            position = next;
            patternPosition += consumed;
            while (patternPosition >= cycle) {
                patternPosition -= cycle;
            }
        }
        patternPosition = AdvancePatternPosition(patternPosition, trailingDistance, cycle);
    }

    private void DrawPatternedPathSegment(
        OfficePoint start,
        OfficePoint end,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double> dashPattern,
        ref double patternPosition) {
        double length = Distance(start.X, start.Y, end.X, end.Y);
        if (!IsFinite(length) || length <= 0D || dashPattern.Count == 0) {
            return;
        }

        dashPattern = NormalizeRasterDashPattern(dashPattern);

        double cycle = 0D;
        for (int i = 0; i < dashPattern.Count; i++) {
            cycle = SaturatingDashCycle(cycle, dashPattern[i]);
            if (cycle == double.MaxValue) break;
        }

        if (!IsFinite(cycle) || cycle <= 0D) {
            return;
        }

        OfficePoint clippedStart = start;
        OfficePoint clippedEnd = end;
        if (!TryClipLineToCanvas(ref clippedStart, ref clippedEnd, thickness, length, out double leadingDistance, out double trailingDistance)) {
            patternPosition = AdvancePatternPosition(patternPosition, length, cycle);
            return;
        }
        patternPosition = AdvancePatternPosition(patternPosition, leadingDistance, cycle);
        length = Distance(clippedStart.X, clippedStart.Y, clippedEnd.X, clippedEnd.Y);
        double position = 0D;
        while (position < length) {
            int patternIndex = 0;
            double patternOffset = patternPosition;
            while (patternIndex < dashPattern.Count && patternOffset >= dashPattern[patternIndex]) {
                patternOffset -= dashPattern[patternIndex];
                patternIndex++;
            }

            if (patternIndex >= dashPattern.Count) {
                patternIndex = 0;
                patternOffset = 0D;
            }

            double segmentRemaining = dashPattern[patternIndex] - patternOffset;
            if (segmentRemaining <= MinimumDashSegmentAdvance) {
                double nextBoundary = 0D;
                for (int index = 0; index <= patternIndex; index++) {
                    nextBoundary = SaturatingDashCycle(nextBoundary, dashPattern[index]);
                }
                patternPosition = nextBoundary >= cycle - MinimumDashSegmentAdvance ? 0D : nextBoundary;
                continue;
            }

            double next = Math.Min(length, position + segmentRemaining);
            double consumed = next - position;
            if (consumed <= MinimumDashSegmentAdvance) {
                break;
            }

            if ((patternIndex & 1) == 0) {
                double startT = position / length;
                double endT = next / length;
                DrawLineSegment(
                    clippedStart.X + ((clippedEnd.X - clippedStart.X) * startT),
                    clippedStart.Y + ((clippedEnd.Y - clippedStart.Y) * startT),
                    clippedStart.X + ((clippedEnd.X - clippedStart.X) * endT),
                    clippedStart.Y + ((clippedEnd.Y - clippedStart.Y) * endT),
                    color,
                    thickness);
            }

            position = next;
            patternPosition += consumed;
            while (patternPosition >= cycle) {
                patternPosition -= cycle;
            }
        }
        patternPosition = AdvancePatternPosition(patternPosition, trailingDistance, cycle);
    }

    private static double SaturatingDashCycle(double left, double right) {
        if (!IsFinite(left) || !IsFinite(right) || left < 0D || right < 0D) return 0D;
        return left > double.MaxValue - right ? double.MaxValue : left + right;
    }

    private static void NormalizeRasterDashLengths(ref double dashLength, ref double gapLength) {
        double smallest = gapLength > 0D ? Math.Min(dashLength, gapLength) : dashLength;
        if (!IsFinite(smallest) || smallest <= 0D || smallest >= MinimumRasterDashLength) return;
        double scale = MinimumRasterDashLength / smallest;
        double normalizedDash = dashLength * scale;
        double normalizedGap = gapLength * scale;
        if (!IsFinite(scale) || !IsFinite(normalizedDash) || !IsFinite(normalizedGap)) {
            dashLength = Math.Max(dashLength, MinimumRasterDashLength);
            if (gapLength > 0D) gapLength = Math.Max(gapLength, MinimumRasterDashLength);
            return;
        }
        dashLength = normalizedDash;
        gapLength = normalizedGap;
    }

    private static IReadOnlyList<double> NormalizeRasterDashPattern(IReadOnlyList<double> pattern) {
        double smallest = double.MaxValue;
        for (int index = 0; index < pattern.Count; index++) {
            double length = pattern[index];
            if (IsFinite(length) && length > 0D) smallest = Math.Min(smallest, length);
        }
        if (smallest == double.MaxValue || smallest >= MinimumRasterDashLength) return pattern;

        double scale = MinimumRasterDashLength / smallest;
        var normalized = new double[pattern.Count];
        for (int index = 0; index < pattern.Count; index++) {
            normalized[index] = pattern[index] * scale;
            if (!IsFinite(scale) || !IsFinite(normalized[index])) {
                for (int fallbackIndex = 0; fallbackIndex < pattern.Count; fallbackIndex++) {
                    normalized[fallbackIndex] = Math.Max(pattern[fallbackIndex], MinimumRasterDashLength);
                }
                break;
            }
        }
        return normalized;
    }
}
