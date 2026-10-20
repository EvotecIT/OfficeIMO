using System;

namespace OfficeIMO.Drawing;

/// <summary>Font-unit line metrics used by document layout, independently of PDF descriptor rounding.</summary>
internal readonly struct OfficeOpenTypeLineMetrics {
    private OfficeOpenTypeLineMetrics(double horizontalAdvance, double ascent, double descent, OfficeOpenTypeDeviceMetrics? deviceMetrics,
        double hheaAscent, double hheaDescent, double hheaLineGap) {
        HorizontalAdvanceRatio = horizontalAdvance;
        AscentRatio = ascent;
        DescentRatio = descent;
        _deviceMetrics = deviceMetrics;
        HheaAscentRatio = hheaAscent;
        HheaDescentRatio = hheaDescent;
        HheaLineGapRatio = hheaLineGap;
    }

    private readonly OfficeOpenTypeDeviceMetrics? _deviceMetrics;
    internal double HorizontalAdvanceRatio { get; }
    internal double AscentRatio { get; }
    internal double DescentRatio { get; }

    // CSS uses the horizontal header independently of Word's OS/2/VDMX policy.
    internal double HheaAscentRatio { get; }
    internal double HheaDescentRatio { get; }
    internal double HheaLineGapRatio { get; }

    // Device metrics affect alignment inside a line, independently of its advance.
    internal void GetDeviceMetrics(double fontSize, double pointsPerPixel, out double ascent, out double descent) {
        double pixels = Math.Max(1D, Math.Round(fontSize / pointsPerPixel, MidpointRounding.AwayFromZero));
        if (pixels <= ushort.MaxValue && _deviceMetrics?.TryGetHeight((int)pixels, out ascent, out descent) == true) return;
        ascent = Math.Round(pixels * AscentRatio, MidpointRounding.AwayFromZero);
        descent = Math.Round(pixels * DescentRatio, MidpointRounding.AwayFromZero);
    }

    internal static OfficeOpenTypeLineMetrics? TryRead(byte[] data) {
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        if (reader == null) return null;
        int advance = reader.Ascender - reader.Descender + reader.LineGap;
        int ascent = reader.Ascender;

        // Some version-zero OS/2 tables end before the Windows metrics. Never
        // read those fields from the following table or beyond the font bytes.
        int descent = Math.Abs((int)reader.Descender);
        if (reader.TryGetTable("OS/2", out int offset, out int length)) {
            if (length >= 78) {
                descent = reader.ReadUInt16(offset + 76);
                ascent = reader.ReadUInt16(offset + 74);
                // Legacy line spacing includes Windows ascent/descent plus any
                // external leading supplied by the horizontal header.
                advance = Math.Max(advance, reader.ReadUInt16(offset + 74) + descent);
            }
            // USE_TYPO_METRICS is defined from OS/2 version four onward. Older
            // versions reserve this bit; a truncated table cannot supply a gap.
            if (length >= 74 && reader.ReadUInt16(offset) >= 4 &&
                (reader.ReadUInt16(offset + 62) & 128) != 0) {
                int typoDescent = reader.ReadInt16(offset + 70);
                int typoAdvance = reader.ReadInt16(offset + 68) - typoDescent + reader.ReadInt16(offset + 72);
                if (typoAdvance > 0) {
                    advance = typoAdvance;
                    ascent = reader.ReadInt16(offset + 68);
                    descent = Math.Abs(typoDescent);
                }
            }
        }
        return advance > 0
            ? new OfficeOpenTypeLineMetrics(advance / (double)reader.UnitsPerEm, ascent / (double)reader.UnitsPerEm,
                descent / (double)reader.UnitsPerEm, OfficeOpenTypeDeviceMetrics.TryRead(reader),
                reader.Ascender / (double)reader.UnitsPerEm, -reader.Descender / (double)reader.UnitsPerEm,
                reader.LineGap / (double)reader.UnitsPerEm)
            : null;
    }
}
