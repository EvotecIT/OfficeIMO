using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing.Binary;

// MS-ODRAW 2.3.6 and 2.4.30 describe native geometry and command consumption.
// Importers own cumulative operation limits and feature-level fidelity reports.
internal static class OfficeArtCustomPathProjector {
    internal static bool HasPath(IReadOnlyList<OfficeArtProperty> properties) =>
        HasArrayElements(Property(properties, 0x0145), true) || HasArrayElements(Property(properties, 0x0146), false);

    private static bool HasArrayElements(OfficeArtProperty? property, bool points) {
        if (property == null || property.Value == 0) return false;
        // A valid complete empty header is absence; malformed nonempty declarations
        // still reach projection so the importer can report their loss.
        return !TryArray(property, points, out _, out int count, out _) || count != 0;
    }

    // Geometry paint controls also apply when an importer falls back to a preset.
    internal static bool IsFillDisabled(IReadOnlyList<OfficeArtProperty> properties) => Disabled(properties, 0x0200);
    internal static bool IsLineDisabled(IReadOnlyList<OfficeArtProperty> properties) => Disabled(properties, 0x0040);

    internal static bool TryProject(IReadOnlyList<OfficeArtProperty> properties,
        double width, double height, Action<int> accountItems, CancellationToken token,
        out OfficeArtCustomPathProjection? projection, out OfficeArtCustomPathFailure failure) {
        projection = null; failure = OfficeArtCustomPathFailure.InvalidData;
        token.ThrowIfCancellationRequested();
        double left = Scalar(properties, 0x0140, 0), top = Scalar(properties, 0x0141, 0);
        double right = Scalar(properties, 0x0142, 21600), bottom = Scalar(properties, 0x0143, 21600);
        if (right <= left || bottom <= top || width < 0 || height < 0 || (width == 0 && height == 0)
            || double.IsNaN(width) || double.IsNaN(height) || double.IsInfinity(width) || double.IsInfinity(height)) {
            failure = OfficeArtCustomPathFailure.InvalidGeometrySpace; return false;
        }
        if (Scalar(properties, 0x0153, int.MinValue) != int.MinValue
            || Scalar(properties, 0x0154, int.MinValue) != int.MinValue) {
            failure = OfficeArtCustomPathFailure.Scaling; return false;
        }
        OfficeArtProperty? vertexProperty = Property(properties, 0x0145);
        if (!TryArray(vertexProperty, true, out OfficeArtComplexData vertices, out int pointCount, out int pointSize)
            || pointCount == 0) return false;
        OfficeArtProperty? segmentProperty = Property(properties, 0x0146);
        OfficeArtComplexData segments = default; int segmentCount = 0;
        if (segmentProperty?.Value > 0 && !TryArray(segmentProperty, false, out segments, out segmentCount, out _)) return false;
        // Charge declared work before allocating decoded arrays or expanding commands.
        accountItems(pointCount + segmentCount);
        var points = new OfficePoint[pointCount];
        OfficeArtGeometryGuides? guides = null;
        for (int index = 0; index < points.Length; index++) {
            token.ThrowIfCancellationRequested();
            int offset = 6 + index * pointSize;
            // Compact Publisher vertices hold unsigned low words; full POINT
            // elements retain signed 32-bit coordinates and guide sentinels.
            int x = pointSize == 8 ? I32(vertices, offset) : U16(vertices, offset);
            int y = pointSize == 8 ? I32(vertices, offset + 4) : U16(vertices, offset + 2);
            if (IsGuide(x) || IsGuide(y)) {
                if (guides == null && !OfficeArtGeometryGuides.TryCreate(properties, width, height,
                    accountItems, token, out guides, out failure)) return false;
                if (IsGuide(x) && !guides!.TryGetValue(x & 0x7F, out x, out failure)
                    || IsGuide(y) && !guides!.TryGetValue(y & 0x7F, out y, out failure)) return false;
            }
            points[index] = new OfficePoint((x - left) / (right - left) * width, (y - top) / (bottom - top) * height);
        }
        var commands = new List<OfficePathCommand>();
        void Add(OfficePathCommand command) {
            token.ThrowIfCancellationRequested(); accountItems(1); commands.Add(command);
        }
        bool noFill = IsFillDisabled(properties), noLine = IsLineDisabled(properties);
        if (segmentCount == 0) {
            int kind = Scalar(properties, 0x0144, 1);
            if (kind < 0 || kind > 3 || (kind >= 2 && (pointCount - 1) % 3 != 0)) return false;
            Add(OfficePathCommand.MoveTo(points[0].X, points[0].Y));
            for (int index = 1; index < points.Length;) {
                if (kind < 2) {
                    Add(OfficePathCommand.LineTo(points[index].X, points[index].Y)); index++;
                } else {
                    Add(Cubic(points[index], points[index + 1], points[index + 2])); index += 3;
                }
            }
            if ((kind & 1) != 0) Add(OfficePathCommand.Close());
        } else {
            int nextPoint = 0; bool started = false, ended = false;
            for (int index = 0; index < segmentCount; index++) {
                token.ThrowIfCancellationRequested();
                ushort word = U16(segments, 6 + index * 2);
                int type = word >> 13, count = word & 0x1FFF;
                if (type == 5) {
                    int escape = (word >> 8) & 31;
                    if (escape is >= 12 and <= 20) {
                        count = word & 0xFF;
                        // Zero-count joint markers add no geometry, including before the first MoveTo.
                        if (count == 0) continue;
                        // Qualified Publisher editing runs share ordinary line/cubic consumption.
                        if (escape is 12 or 14) type = 0;
                        else if (escape is 13 or 15 or 19) type = 1;
                        else { failure = OfficeArtCustomPathFailure.Command; return false; }
                    }
                }
                switch (type) {
                    case 0:
                    case 1:
                        if (!started || count > (pointCount - nextPoint) / (type == 1 ? 3 : 1)) return false;
                        for (int piece = 0; piece < count; piece++) {
                            if (type == 0) {
                                OfficePoint point = points[nextPoint++]; Add(OfficePathCommand.LineTo(point.X, point.Y));
                            } else {
                                Add(Cubic(points[nextPoint], points[nextPoint + 1], points[nextPoint + 2])); nextPoint += 3;
                            }
                        }
                        break;
                    case 2:
                        if (count != 0 || nextPoint == pointCount) return false;
                        OfficePoint start = points[nextPoint++]; Add(OfficePathCommand.MoveTo(start.X, start.Y));
                        started = true; break;
                    case 3:
                        if (count != 1 || !started) return false;
                        Add(OfficePathCommand.Close()); break;
                    case 4:
                        // Move-only paths are valid empty geometry, including picture masks.
                        if (count != 0 || !started) return false;
                        if (index != segmentCount - 1) { failure = OfficeArtCustomPathFailure.PaintGroups; return false; }
                        ended = true; break;
                    case 5:
                        int escape = (word >> 8) & 31;
                        if (escape == 10) noFill = true;
                        else if (escape == 11) noLine = true;
                        else { failure = OfficeArtCustomPathFailure.Command; return false; }
                        break;
                    default:
                        failure = OfficeArtCustomPathFailure.Command; return false;
                }
            }
            if (!ended || nextPoint != pointCount) return false;
        }
        token.ThrowIfCancellationRequested();
        OfficeShape shape = OfficeShape.Path(Math.Max(.01D, width), Math.Max(.01D, height), commands);
        shape.FillRule = OfficeFillRule.NonZero;
        projection = new OfficeArtCustomPathProjection(shape, noFill, noLine, guides != null);
        failure = OfficeArtCustomPathFailure.None; return true;
    }

    private static bool TryArray(OfficeArtProperty? property, bool isPoints,
        out OfficeArtComplexData data, out int count, out int stride) {
        data = default; count = stride = 0;
        if (property?.IsComplex != true) return false;
        OfficeArtComplexData bytes = property.ComplexData;
        if (bytes.Length < 6) return false;
        int elementSize = U16(bytes, 4);
        stride = isPoints && elementSize == 0xFFF0 ? 4 : elementSize;
        count = U16(bytes, 0);
        if (count > U16(bytes, 2) || (isPoints ? stride is not (4 or 8) : stride != 2)
            || 6L + (long)count * stride != bytes.Length) return false;
        data = bytes; return true;
    }

    private static OfficePathCommand Cubic(OfficePoint first, OfficePoint second, OfficePoint end) =>
        OfficePathCommand.CubicBezierTo(first.X, first.Y, second.X, second.Y, end.X, end.Y);
    private static bool IsGuide(int value) => unchecked((uint)value) is >= 0x80000000U and <= 0x8000007FU;
    private static OfficeArtProperty? Property(IReadOnlyList<OfficeArtProperty> properties, ushort id) =>
        properties.LastOrDefault(property => property.PropertyId == id);
    private static int Scalar(IReadOnlyList<OfficeArtProperty> properties, ushort id, int fallback) {
        OfficeArtProperty? property = OfficeArtProperty.GetScalarProperty(properties, id);
        return property == null ? fallback : unchecked((int)property.Value);
    }
    private static bool Disabled(IReadOnlyList<OfficeArtProperty> properties, uint valueBit) =>
        OfficeArtProperty.GetBoolean(properties, 0x017F, valueBit << 16, valueBit) == false;
    private static ushort U16(OfficeArtComplexData bytes, int offset) => unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));
    private static int I32(OfficeArtComplexData bytes, int offset) => bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;
}

internal sealed class OfficeArtCustomPathProjection {
    internal OfficeArtCustomPathProjection(OfficeShape shape, bool noFill, bool noLine, bool usesGuides) {
        Shape = shape; NoFill = noFill; NoLine = noLine; UsesGuides = usesGuides;
    }
    internal OfficeShape Shape { get; }
    internal bool NoFill { get; }
    internal bool NoLine { get; }
    internal bool UsesGuides { get; }
}

internal enum OfficeArtCustomPathFailure {
    None, InvalidData, InvalidGeometrySpace, InvalidGuide, GuideFormula, GuideParameter, Scaling, Command, PaintGroups
}
