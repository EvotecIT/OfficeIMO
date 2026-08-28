using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared structural validation for TrueType composite glyph records.</summary>
internal static class OfficeOpenTypeCompositeGlyph {
    private const ushort HasScale = 0x0008;
    private const ushort HasXAndYScale = 0x0040;
    private const ushort HasTwoByTwo = 0x0080;

    internal static bool HasConflictingTransformFlags(ushort flags) {
        int transformFlags = flags & (HasScale | HasXAndYScale | HasTwoByTwo);
        return transformFlags != 0 && (transformFlags & (transformFlags - 1)) != 0;
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
