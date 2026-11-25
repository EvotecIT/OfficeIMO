using System.Collections.Concurrent;

namespace OfficeIMO.Excel {
    /// <summary>
    /// One style change written to a cell, as one of the model's cell style methods applies it. Operations are shared
    /// up to a bound: the same change with the same arguments is one instance, so recording it allocates nothing, and the
    /// owned save can remember what it derived from each.
    /// </summary>
    internal sealed class ExcelStyleOperation {
        private static readonly ConcurrentDictionary<Key, ExcelStyleOperation> Cache = new();
        // Fonts, colors and number formats are open-ended, so past this many shared operations new ones are not kept;
        // they still apply, and the stylesheet still reuses equal entries.
        private const int CacheLimit = 4096;
        private static int _cached;

        private ExcelStyleOperation(ExcelStyleOperationKind kind) => Kind = kind;

        internal ExcelStyleOperationKind Kind { get; }
        internal bool Flag { get; private set; }
        /// <summary>A token or text argument: an underline or alignment token, a font name or a number format.</summary>
        internal string? Text { get; private set; }
        internal double Number { get; private set; }
        internal ExcelStyleColor? Color { get; private set; }
        internal ExcelCellBorders Borders { get; private set; }
        /// <summary>The argument the caller passed, for applying the change through the model when the package opens.</summary>
        private object? Source { get; set; }

        /// <summary>The automatic date format a date value takes.</summary>
        internal static readonly ExcelStyleOperation Date = new(ExcelStyleOperationKind.Date);

        private readonly record struct Key(ExcelStyleOperationKind Kind, bool Flag, string? Text, double Number, string? ColorAttribute,
            string? ColorValue, ExcelCellBorders Borders);

        private static ExcelStyleOperation Get(ExcelStyleOperationKind kind, bool flag = false, string? text = null, double number = 0,
            ExcelStyleColor? color = null, ExcelCellBorders borders = ExcelCellBorders.None, object? source = null) {
            var key = new Key(kind, flag, text, number, color?.Attribute, color?.Value, borders);
            if (Cache.TryGetValue(key, out ExcelStyleOperation? known)) return known;
            var operation = new ExcelStyleOperation(kind) { Flag = flag, Text = text, Number = number, Color = color, Borders = borders, Source = source };
            if (Volatile.Read(ref _cached) >= CacheLimit) return operation;
            if (Cache.TryAdd(key, operation)) {
                Interlocked.Increment(ref _cached);
                return operation;
            }
            return Cache.TryGetValue(key, out known) ? known : operation;
        }

        internal static ExcelStyleOperation Flagged(ExcelStyleOperationKind kind, bool flag) => Get(kind, flag);

        internal static ExcelStyleOperation Underline(bool underline) => Get(ExcelStyleOperationKind.Underline, text: underline ? string.Empty : null, source: underline);

        internal static ExcelStyleOperation Underline(ExcelUnderlineStyle style) =>
            Get(ExcelStyleOperationKind.Underline, text: style == ExcelUnderlineStyle.None ? null : style.ToOpenXmlToken(), source: style);

        internal static ExcelStyleOperation VerticalTextAlignment(ExcelVerticalTextAlignment alignment) =>
            Get(ExcelStyleOperationKind.VerticalTextAlignment, text: alignment.ToOpenXmlToken(), source: alignment);

        internal static ExcelStyleOperation FontName(string fontName) => Get(ExcelStyleOperationKind.FontName, text: fontName);

        internal static ExcelStyleOperation FontSize(OfficeLength size) => Get(ExcelStyleOperationKind.FontSize, number: size.ToPoints(), source: size);

        internal static ExcelStyleOperation FontColor(ExcelStyleColor color, OfficeColor source) => Get(ExcelStyleOperationKind.FontColor, color: color, source: source);

        internal static ExcelStyleOperation Fill(ExcelStyleColor color, object source) => Get(ExcelStyleOperationKind.Fill, color: color, source: source);

        internal static ExcelStyleOperation Border(ExcelCellBorders borders, ExcelBorderStyle style, ExcelStyleColor? color, OfficeColor? source) =>
            Get(ExcelStyleOperationKind.Border, text: style.ToOpenXmlToken(), color: color, borders: borders, source: (style, source));

        internal static ExcelStyleOperation NumberFormat(string code) => Get(ExcelStyleOperationKind.NumberFormat, text: code);

        internal static ExcelStyleOperation HorizontalAlignment(ExcelHorizontalAlignment alignment) =>
            Get(ExcelStyleOperationKind.HorizontalAlignment, text: alignment.ToOpenXmlToken(), source: alignment);

        internal static ExcelStyleOperation VerticalAlignment(ExcelVerticalAlignment alignment) =>
            Get(ExcelStyleOperationKind.VerticalAlignment, text: alignment.ToOpenXmlToken(), source: alignment);

        internal static ExcelStyleOperation TextRotation(int rotation) => Get(ExcelStyleOperationKind.TextRotation, number: rotation);

        /// <summary>Applies the change through the model, as the style method it came from would have.</summary>
        internal void Replay(ExcelSheet sheet, int row, int column) {
            switch (Kind) {
                case ExcelStyleOperationKind.Bold: sheet.CellBold(row, column, Flag); break;
                case ExcelStyleOperationKind.Italic: sheet.CellItalic(row, column, Flag); break;
                case ExcelStyleOperationKind.Strike: sheet.CellStrikethrough(row, column, Flag); break;
                case ExcelStyleOperationKind.WrapText: sheet.CellWrapText(row, column, Flag); break;
                case ExcelStyleOperationKind.ShrinkToFit: sheet.CellShrinkToFit(row, column, Flag); break;
                case ExcelStyleOperationKind.Underline:
                    if (Source is ExcelUnderlineStyle style) sheet.CellUnderline(row, column, style);
                    else sheet.CellUnderline(row, column, Text != null);
                    break;
                case ExcelStyleOperationKind.VerticalTextAlignment:
                    sheet.CellVerticalTextAlignment(row, column, (ExcelVerticalTextAlignment)Source!);
                    break;
                case ExcelStyleOperationKind.FontName: sheet.CellFontName(row, column, Text!); break;
                case ExcelStyleOperationKind.FontSize: sheet.CellFontSize(row, column, (OfficeLength)Source!); break;
                case ExcelStyleOperationKind.FontColor: sheet.CellFontColor(row, column, (OfficeColor)Source!); break;
                case ExcelStyleOperationKind.Fill:
                    if (Source is string hex) sheet.CellBackground(row, column, hex);
                    else sheet.CellBackground(row, column, (OfficeColor)Source!);
                    break;
                case ExcelStyleOperationKind.Border:
                    var (borderStyle, borderColor) = ((ExcelBorderStyle, OfficeColor?))Source!;
                    sheet.CellBorder(row, column, Borders, borderStyle, borderColor);
                    break;
                case ExcelStyleOperationKind.NumberFormat: sheet.FormatCell(row, column, Text!); break;
                case ExcelStyleOperationKind.HorizontalAlignment: sheet.CellAlign(row, column, (ExcelHorizontalAlignment)Source!); break;
                case ExcelStyleOperationKind.VerticalAlignment: sheet.CellVerticalAlign(row, column, (ExcelVerticalAlignment)Source!); break;
                case ExcelStyleOperationKind.TextRotation: sheet.CellTextRotation(row, column, (int)Number); break;
                default: throw new InvalidOperationException("A recorded style change has no model method.");
            }
        }
    }
}
