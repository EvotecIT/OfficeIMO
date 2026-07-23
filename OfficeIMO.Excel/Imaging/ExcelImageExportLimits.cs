using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Excel {
    internal static class ExcelImageExportLimits {
        internal const int MaximumAnchorSpanCells = 10_000;
        internal const int MaximumAnchorOffsetPixels = 100_000;
        internal const int MaximumAnchorExtentPixels = 16_384;
        internal const double MaximumStrokeWidth = 64D;

        internal static int ClampOffsetPixels(int value) =>
            Math.Max(-MaximumAnchorOffsetPixels, Math.Min(MaximumAnchorOffsetPixels, value));

        internal static int ClampAbsoluteOffsetPixels(int value) =>
            Math.Max(0, Math.Min(MaximumAnchorOffsetPixels, value));

        internal static int ClampExtentPixels(int value) =>
            Math.Max(0, Math.Min(MaximumAnchorExtentPixels, value));

        internal static int EmuOffsetDifferencePixels(long fromOffset, long toOffset) {
            double pixels = Math.Round(((double)toOffset - fromOffset) / 9525D);
            if (pixels <= -MaximumAnchorOffsetPixels) return -MaximumAnchorOffsetPixels;
            if (pixels >= MaximumAnchorOffsetPixels) return MaximumAnchorOffsetPixels;
            return (int)pixels;
        }

        internal static double ClampStrokeWidth(double value) =>
            Math.Max(0D, Math.Min(MaximumStrokeWidth, value));

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

        internal static bool TryReadSourceImageBytes(Stream source, out byte[] bytes) =>
            TryReadSourceImageBytes(source, ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes, out bytes);

        internal static bool TryReadSourceImageBytes(Stream source, long remainingAggregateBytes, out byte[] bytes) {
            if (remainingAggregateBytes <= 0L) {
                bytes = Array.Empty<byte>();
                return false;
            }

            try {
                bytes = OfficeStreamReader.ReadAllBytes(source, remainingAggregateBytes);
                return bytes.Length > 0;
            } catch (InvalidDataException) {
                bytes = Array.Empty<byte>();
                return false;
            }
        }
    }
}
