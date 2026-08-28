using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared structural validation for TrueType composite glyph records.</summary>
internal static class OfficeOpenTypeCompositeGlyph {
    private const ushort HasScale = 0x0008;
    private const ushort HasXAndYScale = 0x0040;
    private const ushort HasTwoByTwo = 0x0080;
    private const ushort ScaledComponentOffset = 0x0800;
    private const ushort UnscaledComponentOffset = 0x1000;

    internal static bool HasConflictingTransformFlags(ushort flags) {
        int transformFlags = flags & (HasScale | HasXAndYScale | HasTwoByTwo);
        return transformFlags != 0 && (transformFlags & (transformFlags - 1)) != 0;
    }

    internal static bool HasConflictingOffsetFlags(ushort flags) =>
        (flags & ScaledComponentOffset) != 0 && (flags & UnscaledComponentOffset) != 0;

    internal static OfficePoint ResolveXyOffset(
        ushort flags,
        double xx,
        double xy,
        double yx,
        double yy,
        double x,
        double y) {
        if (HasConflictingOffsetFlags(flags)) {
            throw new System.ArgumentException("Composite glyph offset flags conflict.", nameof(flags));
        }
        return (flags & ScaledComponentOffset) != 0
            ? new OfficePoint((xx * x) + (xy * y), (yx * x) + (yy * y))
            : new OfficePoint(x, y);
    }

    internal static bool TryResolvePointAttachment(
        IReadOnlyList<OfficePoint> parentPoints,
        int parentPointIndex,
        IReadOnlyList<OfficePoint> componentPoints,
        int componentPointIndex,
        out OfficePoint translation) {
        if ((uint)parentPointIndex >= (uint)parentPoints.Count ||
            (uint)componentPointIndex >= (uint)componentPoints.Count) {
            translation = default;
            return false;
        }

        OfficePoint parent = parentPoints[parentPointIndex];
        OfficePoint component = componentPoints[componentPointIndex];
        translation = new OfficePoint(parent.X - component.X, parent.Y - component.Y);
        return true;
    }

    internal static OfficePoint ApplyComponentVariation(OfficePoint attachmentTranslation, OfficePoint variationTranslation) =>
        new OfficePoint(
            attachmentTranslation.X + variationTranslation.X,
            attachmentTranslation.Y + variationTranslation.Y);
}
