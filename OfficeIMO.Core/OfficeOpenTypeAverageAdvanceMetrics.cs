namespace OfficeIMO.Drawing;

/// <summary>Optional source-font average advance in design units, separate from shaped glyph measurements.</summary>
internal readonly struct OfficeOpenTypeAverageAdvanceMetrics {
    private OfficeOpenTypeAverageAdvanceMetrics(int averageAdvance, int unitsPerEm) {
        AverageAdvance = averageAdvance;
        UnitsPerEm = unitsPerEm;
    }

    /// <summary>Positive OS/2 xAvgCharWidth in the source font's design units.</summary>
    internal int AverageAdvance { get; }

    /// <summary>Source head-table units per em used to interpret the average advance.</summary>
    internal int UnitsPerEm { get; }

    /// <summary>Reads optional metrics without rejecting a font when the source field is unavailable.</summary>
    internal static OfficeOpenTypeAverageAdvanceMetrics? TryRead(byte[]? data) =>
        data == null ? null : TryRead(OfficeOpenTypeReader.TryCreate(data));

    /// <summary>Reuses a validated reader without reparsing the font or changing its glyph advances.</summary>
    internal static OfficeOpenTypeAverageAdvanceMetrics? TryRead(OfficeOpenTypeReader? reader) {
        // Bound the read by the declared table, even when later backing bytes exist.
        if (reader == null || !reader.TryGetTable("OS/2", out int offset, out int length) || length < 4)
            return null;
        int averageAdvance = reader.ReadInt16(offset + 2);
        return averageAdvance > 0
            ? new OfficeOpenTypeAverageAdvanceMetrics(averageAdvance, reader.UnitsPerEm)
            : null;
    }
}
