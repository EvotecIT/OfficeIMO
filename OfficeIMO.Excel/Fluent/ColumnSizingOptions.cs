using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Options for clamping column widths and enabling wrapping by header.
    /// Widths use Excel's character-based units (approx. number of '0' glyphs).
    /// </summary>
    public sealed class ColumnSizingOptions {
        public double ShortWidth { get; set; } = 16;     // Status/Alg/Hash
        public double NumericWidth { get; set; } = 10;   // Depth/Count/Days
        public double MediumWidth { get; set; } = 28;    // Provider/Target/Reason
        public double LongWidth { get; set; } = 56;      // Evidence/Record/URL/Summary

        public HashSet<string> ShortHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> NumericHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> MediumHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> LongHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> WrapHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        // Explicit overrides
        public Dictionary<string, double> WidthByHeader { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}

