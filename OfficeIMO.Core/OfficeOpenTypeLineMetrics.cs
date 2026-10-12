using System;

namespace OfficeIMO.Drawing;

/// <summary>Font-unit line metrics used by document layout, independently of PDF descriptor rounding.</summary>
internal readonly struct OfficeOpenTypeLineMetrics {
    private OfficeOpenTypeLineMetrics(double horizontalAdvance, double descent) {
        HorizontalAdvanceRatio = horizontalAdvance;
        DescentRatio = descent;
    }

    internal double HorizontalAdvanceRatio { get; }
    internal double DescentRatio { get; }

    internal static OfficeOpenTypeLineMetrics? TryRead(byte[] data) {
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        if (reader == null) return null;
        int advance = reader.Ascender - reader.Descender + reader.LineGap;

        // Some version-zero OS/2 tables end before the Windows metrics. Never
        // read those fields from the following table or beyond the font bytes.
        int descent = Math.Abs((int)reader.Descender);
        if (reader.TryGetTable("OS/2", out int offset, out int length)) {
            if (length >= 78) {
                descent = reader.ReadUInt16(offset + 76);
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
                    descent = Math.Abs(typoDescent);
                }
            }
        }
        return advance > 0
            ? new OfficeOpenTypeLineMetrics(advance / (double)reader.UnitsPerEm, descent / (double)reader.UnitsPerEm)
            : null;
    }
}
