namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a parsed legacy XLS Font record.
    /// </summary>
    public sealed class LegacyXlsFont {
        /// <summary>
        /// Creates a parsed legacy XLS font.
        /// </summary>
        /// <param name="fontIndex">BIFF FontIndex value, which skips the reserved value 4.</param>
        /// <param name="name">Font family name.</param>
        /// <param name="size">Font size in points, when present.</param>
        /// <param name="colorIndex">Legacy IcvFont color index.</param>
        /// <param name="bold">Whether the font weight is bold.</param>
        /// <param name="italic">Whether the font is italic.</param>
        /// <param name="underline">Whether the font uses an underline style.</param>
        /// <param name="strikeout">Whether the font uses strikethrough.</param>
        /// <param name="underlineStyle">BIFF underline style byte.</param>
        /// <param name="escapement">Legacy superscript or subscript positioning.</param>
        /// <param name="family">BIFF font family classification byte.</param>
        /// <param name="characterSet">BIFF font character set byte.</param>
        /// <param name="outline">Whether the font uses the BIFF outline flag.</param>
        /// <param name="shadow">Whether the font uses the BIFF shadow flag.</param>
        /// <param name="condense">Whether the font uses the BIFF condense flag.</param>
        /// <param name="extend">Whether the font uses the BIFF extend flag.</param>
        public LegacyXlsFont(
            ushort fontIndex,
            string? name,
            double? size,
            ushort colorIndex,
            bool bold,
            bool italic,
            bool underline,
            bool strikeout,
            byte underlineStyle = 0,
            LegacyXlsFontEscapement escapement = LegacyXlsFontEscapement.None,
            byte family = 0,
            byte characterSet = 1,
            bool outline = false,
            bool shadow = false,
            bool condense = false,
            bool extend = false) {
            FontIndex = fontIndex;
            Name = name;
            Size = size;
            ColorIndex = colorIndex;
            Bold = bold;
            Italic = italic;
            UnderlineStyle = underlineStyle == 0 && underline ? (byte)1 : underlineStyle;
            Strikeout = strikeout;
            Escapement = escapement;
            Family = family;
            CharacterSet = characterSet;
            Outline = outline;
            Shadow = shadow;
            Condense = condense;
            Extend = extend;
        }

        /// <summary>
        /// Gets the BIFF FontIndex value, which skips the reserved value 4.
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
        public bool Underline => UnderlineStyle != 0;

        /// <summary>
        /// Gets the BIFF underline style byte.
        /// </summary>
        public byte UnderlineStyle { get; }

        /// <summary>
        /// Gets whether the font uses strikethrough.
        /// </summary>
        public bool Strikeout { get; }

        /// <summary>
        /// Gets superscript or subscript positioning for the font.
        /// </summary>
        public LegacyXlsFontEscapement Escapement { get; }

        /// <summary>
        /// Gets the BIFF font family classification byte.
        /// </summary>
        public byte Family { get; }

        /// <summary>
        /// Gets the BIFF font character set byte.
        /// </summary>
        public byte CharacterSet { get; }

        /// <summary>
        /// Gets whether the font uses the BIFF outline flag.
        /// </summary>
        public bool Outline { get; }

        /// <summary>
        /// Gets whether the font uses the BIFF shadow flag.
        /// </summary>
        public bool Shadow { get; }

        /// <summary>
        /// Gets whether the font uses the BIFF condense flag.
        /// </summary>
        public bool Condense { get; }

        /// <summary>
        /// Gets whether the font uses the BIFF extend flag.
        /// </summary>
        public bool Extend { get; }
    }
}
