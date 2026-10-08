using System;

namespace OfficeIMO.Drawing;

/// <summary>Font-unit line metrics used by document layout, independently of PDF descriptor rounding.</summary>
internal readonly struct OfficeOpenTypeLineMetrics {
    private OfficeOpenTypeLineMetrics(double horizontalAdvance, double windowsDescent) {
        HorizontalAdvanceRatio = horizontalAdvance;
        WindowsDescentRatio = windowsDescent;
    }

    internal double HorizontalAdvanceRatio { get; }
    internal double WindowsDescentRatio { get; }

    internal static OfficeOpenTypeLineMetrics? TryRead(byte[] data) {
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        if (reader == null) return null;
        double advance = (reader.Ascender - reader.Descender + reader.LineGap) / (double)reader.UnitsPerEm;
        if (advance <= 0D) return null;

        // Some version-zero OS/2 tables end before the Windows metrics. Never
        // read those fields from the following table or beyond the font bytes.
        int descent = Math.Abs((int)reader.Descender);
        if (reader.TryGetTable("OS/2", out int offset, out int length) && length >= 78)
            descent = reader.ReadUInt16(offset + 76);
        return new OfficeOpenTypeLineMetrics(advance, descent / (double)reader.UnitsPerEm);
    }
}
