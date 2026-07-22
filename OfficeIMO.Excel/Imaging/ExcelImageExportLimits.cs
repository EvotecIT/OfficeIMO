namespace OfficeIMO.Excel {
    internal static class ExcelImageExportLimits {
        internal const int MaximumAnchorSpanCells = 10_000;
        internal const int MaximumAnchorOffsetPixels = 100_000;
        internal const int MaximumAnchorExtentPixels = 16_384;

        internal static int ClampOffsetPixels(int value) =>
            Math.Max(-MaximumAnchorOffsetPixels, Math.Min(MaximumAnchorOffsetPixels, value));

        internal static int ClampAbsoluteOffsetPixels(int value) =>
            Math.Max(0, Math.Min(MaximumAnchorOffsetPixels, value));

        internal static int ClampExtentPixels(int value) =>
            Math.Max(0, Math.Min(MaximumAnchorExtentPixels, value));

        internal static int SaturatingAddExtent(int sizePixels, int offsetPixels) {
            long total = Math.Max(1, sizePixels) + Math.Max(0, (long)offsetPixels);
            return (int)Math.Min(MaximumAnchorExtentPixels, total);
        }

        internal static int SaturatingAddAbsoluteOffset(int offsetPixels, int extentPixels) {
            long total = Math.Max(0, (long)offsetPixels) + Math.Max(1, (long)extentPixels);
            return (int)Math.Min(MaximumAnchorOffsetPixels, total);
        }

        internal static bool IsValidMarkerSpan(int from, int to, int maximumMarker) =>
            from >= 0 &&
            to >= from &&
            from < maximumMarker &&
            to < maximumMarker &&
            to - from <= MaximumAnchorSpanCells;
    }
}
