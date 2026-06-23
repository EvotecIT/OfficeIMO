namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a parsed legacy XLS Font record.
    /// </summary>
    public sealed class LegacyXlsFont {
        /// <summary>
        /// Creates a parsed legacy XLS font.
        /// </summary>
        /// <param name="fontIndex">Zero-based position in the Font record collection.</param>
        /// <param name="name">Font family name.</param>
        /// <param name="size">Font size in points, when present.</param>
        /// <param name="colorIndex">Legacy IcvFont color index.</param>
        /// <param name="bold">Whether the font weight is bold.</param>
        /// <param name="italic">Whether the font is italic.</param>
        /// <param name="underline">Whether the font uses an underline style.</param>
        /// <param name="strikeout">Whether the font uses strikethrough.</param>
        /// <param name="escapement">Legacy superscript or subscript positioning.</param>
        public LegacyXlsFont(
            ushort fontIndex,
            string? name,
            double? size,
            ushort colorIndex,
            bool bold,
            bool italic,
            bool underline,
            bool strikeout,
            LegacyXlsFontEscapement escapement = LegacyXlsFontEscapement.None) {
            FontIndex = fontIndex;
            Name = name;
            Size = size;
            ColorIndex = colorIndex;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strikeout = strikeout;
            Escapement = escapement;
        }

        /// <summary>
        /// Gets the zero-based position in the Font record collection.
        /// </summary>
        public ushort FontIndex { get; }

        /// <summary>
        /// Gets the font family name.
        /// </summary>
        public string? Name { get; }

        /// <summary>
        /// Gets the font size in points.
        /// </summary>
        public double? Size { get; }

        /// <summary>
        /// Gets the legacy IcvFont color index.
        /// </summary>
        public ushort ColorIndex { get; }

        /// <summary>
        /// Gets whether the font weight is bold.
        /// </summary>
        public bool Bold { get; }

        /// <summary>
        /// Gets whether the font is italic.
        /// </summary>
        public bool Italic { get; }

        /// <summary>
        /// Gets whether the font uses an underline style.
        /// </summary>
        public bool Underline { get; }

        /// <summary>
        /// Gets whether the font uses strikethrough.
        /// </summary>
        public bool Strikeout { get; }

        /// <summary>
        /// Gets superscript or subscript positioning for the font.
        /// </summary>
        public LegacyXlsFontEscapement Escapement { get; }
    }
}
