namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a custom client color palette decoded from a legacy XLS chart ClrtClient record.
    /// </summary>
    public sealed class LegacyXlsChartClientColorPalette {
        internal LegacyXlsChartClientColorPalette(short declaredColorCount, IReadOnlyList<string> colors) {
            DeclaredColorCount = declaredColorCount;
            Colors = colors ?? throw new ArgumentNullException(nameof(colors));
        }

        /// <summary>Gets the color count declared by the record.</summary>
        public short DeclaredColorCount { get; }

        /// <summary>Gets decoded LongRGB colors in record order.</summary>
        public IReadOnlyList<string> Colors { get; }

        /// <summary>Gets the number of colors decoded from the available payload.</summary>
        public int DecodedColorCount => Colors.Count;

        /// <summary>Gets whether the decoded color count matches the declared count.</summary>
        public bool HasCompleteColorList => DeclaredColorCount == Colors.Count;

        /// <summary>Gets whether the palette declares the expected three colors.</summary>
        public bool HasExpectedColorCount => DeclaredColorCount == 3 && Colors.Count == 3;

        /// <summary>Gets the foreground color, when present.</summary>
        public string? ForegroundColor => Colors.Count > 0 ? Colors[0] : null;

        /// <summary>Gets the background color, when present.</summary>
        public string? BackgroundColor => Colors.Count > 1 ? Colors[1] : null;

        /// <summary>Gets the neutral color, when present.</summary>
        public string? NeutralColor => Colors.Count > 2 ? Colors[2] : null;
    }
}
