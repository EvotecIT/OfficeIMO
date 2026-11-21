namespace OfficeIMO.Drawing;

/// <summary>Optional decoration metrics in font design units, independent of rendering policy.</summary>
internal readonly struct OfficeOpenTypeDecorationMetrics {
    private OfficeOpenTypeDecorationMetrics(Metric? strikethrough, Metric? underline, ScriptMetric? subscript, ScriptMetric? superscript) {
        Strikethrough = strikethrough;
        Underline = underline;
        Subscript = subscript;
        Superscript = superscript;
    }

    internal Metric? Strikethrough { get; }
    internal Metric? Underline { get; }
    /// <summary>The OS/2 subscript size and downward offset, as fractions of the em.</summary>
    internal ScriptMetric? Subscript { get; }
    /// <summary>The OS/2 superscript size and upward offset, as fractions of the em.</summary>
    internal ScriptMetric? Superscript { get; }

    internal static OfficeOpenTypeDecorationMetrics? TryRead(byte[]? data) {
        if (data == null) return null;
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        if (reader == null) return null;

        Metric? strikethrough = null;
        Metric? underline = null;
        ScriptMetric? subscript = null;
        ScriptMetric? superscript = null;
        // These fields are present in every OS/2 version. The declared table
        // length, rather than the remaining font bytes, bounds both reads.
        if (reader.TryGetTable("OS/2", out int offset, out int length) && length >= 30) {
            int thickness = reader.ReadInt16(offset + 26);
            if (thickness > 0) {
                strikethrough = new Metric(reader.ReadInt16(offset + 28), thickness, reader.UnitsPerEm);
            }
            subscript = ScriptMetric.TryCreate(reader.ReadInt16(offset + 12), reader.ReadInt16(offset + 16), reader.UnitsPerEm);
            superscript = ScriptMetric.TryCreate(reader.ReadInt16(offset + 20), reader.ReadInt16(offset + 24), reader.UnitsPerEm);
        }
        if (reader.TryGetTable("post", out offset, out length) && length >= 12) {
            int thickness = reader.ReadInt16(offset + 10);
            if (thickness > 0) {
                underline = new Metric(reader.ReadInt16(offset + 8), thickness, reader.UnitsPerEm);
            }
        }

        return strikethrough.HasValue || underline.HasValue || subscript.HasValue || superscript.HasValue
            ? new OfficeOpenTypeDecorationMetrics(strikethrough, underline, subscript, superscript)
            : null;
    }

    /// <summary>A script size and vertical offset normalized by units per em.</summary>
    internal readonly struct ScriptMetric {
        private ScriptMetric(double sizeRatio, double offsetRatio) {
            SizeRatio = sizeRatio;
            OffsetRatio = offsetRatio;
        }

        internal double SizeRatio { get; }
        /// <summary>The offset's magnitude; fonts store subscript offsets with either sign.</summary>
        internal double OffsetRatio { get; }

        internal static ScriptMetric? TryCreate(int size, int offset, int unitsPerEm) {
            double sizeRatio = size / (double)unitsPerEm, offsetRatio = Math.Abs(offset) / (double)unitsPerEm;
            return sizeRatio > 0D && sizeRatio <= 1D && offsetRatio <= 1D ? new ScriptMetric(sizeRatio, offsetRatio) : null;
        }
    }

    /// <summary>Signed font position and positive stroke thickness normalized by units per em.</summary>
    internal readonly struct Metric {
        internal Metric(int position, int thickness, int unitsPerEm) {
            PositionRatio = position / (double)unitsPerEm;
            ThicknessRatio = thickness / (double)unitsPerEm;
        }

        internal double PositionRatio { get; }
        internal double ThicknessRatio { get; }

        /// <summary>Quantizes to device pixels and returns both measurements in points.</summary>
        /// <remarks>The position retains the font's sign. This does not choose a rendering origin or device resolution.</remarks>
        internal void GetDeviceMetrics(double fontSize, double pointsPerPixel,
            out double positionPoints, out double thicknessPoints) {
            if (!IsPositiveFinite(fontSize)) throw new ArgumentOutOfRangeException(nameof(fontSize));
            if (!IsPositiveFinite(pointsPerPixel)) throw new ArgumentOutOfRangeException(nameof(pointsPerPixel));

            double pixels = Math.Round(fontSize / pointsPerPixel, MidpointRounding.AwayFromZero);
            double position = Math.Round(pixels * PositionRatio, MidpointRounding.AwayFromZero);
            double thickness = Math.Round(pixels * ThicknessRatio, MidpointRounding.AwayFromZero);
            if (ThicknessRatio > 0D) thickness = Math.Max(1D, thickness);
            positionPoints = position * pointsPerPixel;
            thicknessPoints = thickness * pointsPerPixel;
            if (double.IsNaN(positionPoints) || double.IsInfinity(positionPoints) ||
                double.IsNaN(thicknessPoints) || double.IsInfinity(thicknessPoints)) {
                throw new ArgumentOutOfRangeException(nameof(fontSize));
            }
        }

        private static bool IsPositiveFinite(double value) => value > 0D && !double.IsInfinity(value);
    }
}
