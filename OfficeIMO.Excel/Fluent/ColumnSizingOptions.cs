using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Options for clamping column widths and enabling wrapping by header.
    /// Widths use Excel's character-based units (approx. number of '0' glyphs).
    /// </summary>
    public sealed class ColumnSizingOptions {
        /// <summary>
        /// Width in characters used for headers registered in <see cref="ShortHeaders"/>.
        /// </summary>
        public double ShortWidth { get; set; } = 16;     // Status/Alg/Hash

        /// <summary>
        /// Width in characters used for headers registered in <see cref="NumericHeaders"/>.
        /// </summary>
        public double NumericWidth { get; set; } = 10;   // Depth/Count/Days

        /// <summary>
        /// Width in characters used for headers registered in <see cref="MediumHeaders"/>.
        /// </summary>
        public double MediumWidth { get; set; } = 28;    // Provider/Target/Reason

        /// <summary>
        /// Width in characters used for headers registered in <see cref="LongHeaders"/>.
        /// </summary>
        public double LongWidth { get; set; } = 56;      // Evidence/Record/URL/Summary

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="ShortWidth"/>.
        /// </summary>
        public HashSet<string> ShortHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="NumericWidth"/>.
        /// </summary>
        public HashSet<string> NumericHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="MediumWidth"/>.
        /// </summary>
        public HashSet<string> MediumHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="LongWidth"/>.
        /// </summary>
        public HashSet<string> LongHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should be wrapped regardless of width.
        /// </summary>
        public HashSet<string> WrapHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Explicit column width overrides keyed by header name. Values are Excel character widths.
        /// </summary>
        public Dictionary<string, double> WidthByHeader { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}

