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
}
