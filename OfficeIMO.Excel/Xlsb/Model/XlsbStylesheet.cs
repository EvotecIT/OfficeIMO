namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents the reusable formatting collections referenced by XLSB cells.</summary>
    internal sealed class XlsbStylesheet {
        private readonly Dictionary<ushort, string> _numberFormats = new Dictionary<ushort, string>();
        private readonly List<XlsbFont> _fonts = new List<XlsbFont>();
        private readonly List<XlsbFill> _fills = new List<XlsbFill>();
        private readonly List<XlsbBorder> _borders = new List<XlsbBorder>();
        private readonly List<XlsbCellFormat> _cellStyleFormats = new List<XlsbCellFormat>();
        private readonly List<XlsbCellFormat> _cellFormats = new List<XlsbCellFormat>();

        internal IReadOnlyDictionary<ushort, string> NumberFormats => _numberFormats;

        internal IReadOnlyList<XlsbFont> Fonts => _fonts;

        internal IReadOnlyList<XlsbFill> Fills => _fills;

        internal IReadOnlyList<XlsbBorder> Borders => _borders;

        internal IReadOnlyList<XlsbCellFormat> CellStyleFormats => _cellStyleFormats;

        internal IReadOnlyList<XlsbCellFormat> CellFormats => _cellFormats;

        internal void AddNumberFormat(ushort id, string code) => _numberFormats.Add(id, code);

        internal void AddFont(XlsbFont font) => _fonts.Add(font);

        internal void AddFill(XlsbFill fill) => _fills.Add(fill);

        internal void AddBorder(XlsbBorder border) => _borders.Add(border);

        internal void AddCellStyleFormat(XlsbCellFormat format) => _cellStyleFormats.Add(format);

        internal void AddCellFormat(XlsbCellFormat format) => _cellFormats.Add(format);
    }

    internal sealed class XlsbColor {
        internal XlsbColor(byte type, byte index, short tint, byte red, byte green, byte blue, byte alpha) {
            Type = type;
            Index = index;
            Tint = tint;
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
        }

        internal byte Type { get; }
        internal byte Index { get; }
        internal short Tint { get; }
        internal byte Red { get; }
        internal byte Green { get; }
        internal byte Blue { get; }
        internal byte Alpha { get; }
    }

    internal sealed class XlsbFont {
        internal ushort HeightTwips { get; set; }
        internal ushort Flags { get; set; }
        internal ushort Weight { get; set; }
        internal ushort Script { get; set; }
        internal byte Underline { get; set; }
        internal byte Family { get; set; }
        internal byte CharacterSet { get; set; }
        internal XlsbColor? Color { get; set; }
        internal byte Scheme { get; set; }
        internal string Name { get; set; } = string.Empty;
    }

    internal sealed class XlsbFill {
        internal uint Pattern { get; set; }
        internal XlsbColor? Foreground { get; set; }
        internal XlsbColor? Background { get; set; }
        internal int GradientType { get; set; }
        internal uint GradientStopCount { get; set; }
    }

    internal sealed class XlsbBorder {
        internal bool DiagonalDown { get; set; }
        internal bool DiagonalUp { get; set; }
        internal XlsbBorderSide Top { get; set; } = XlsbBorderSide.None;
        internal XlsbBorderSide Bottom { get; set; } = XlsbBorderSide.None;
        internal XlsbBorderSide Left { get; set; } = XlsbBorderSide.None;
        internal XlsbBorderSide Right { get; set; } = XlsbBorderSide.None;
        internal XlsbBorderSide Diagonal { get; set; } = XlsbBorderSide.None;
    }

    internal sealed class XlsbBorderSide {
        internal static XlsbBorderSide None { get; } = new XlsbBorderSide(0, null);

        internal XlsbBorderSide(byte style, XlsbColor? color) {
            Style = style;
            Color = color;
        }

        internal byte Style { get; }
        internal XlsbColor? Color { get; }
    }

    internal sealed class XlsbCellFormat {
        internal ushort ParentFormatId { get; set; }
        internal ushort NumberFormatId { get; set; }
        internal ushort FontId { get; set; }
        internal ushort FillId { get; set; }
        internal ushort BorderId { get; set; }
        internal byte TextRotation { get; set; }
        internal byte Indent { get; set; }
        internal byte HorizontalAlignment { get; set; }
        internal byte VerticalAlignment { get; set; }
        internal bool WrapText { get; set; }
        internal bool JustifyLastLine { get; set; }
        internal bool ShrinkToFit { get; set; }
        internal bool Merged { get; set; }
        internal byte ReadingOrder { get; set; }
        internal bool Locked { get; set; }
        internal bool Hidden { get; set; }
        internal bool PivotButton { get; set; }
        internal bool QuotePrefix { get; set; }
        internal byte ApplyFlags { get; set; }
    }
}
