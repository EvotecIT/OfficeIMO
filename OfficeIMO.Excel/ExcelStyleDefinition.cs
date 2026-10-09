using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Declares a complete cell style based on the workbook's Normal style, without requiring a source cell.
    /// Definitions are copied when a named style or tabular export is created.
    /// </summary>
    public sealed class ExcelStyleDefinition {
        /// <summary>Displays text in bold.</summary>
        public bool Bold { get; set; }
        /// <summary>Displays text in italics.</summary>
        public bool Italic { get; set; }
        /// <summary>Optional underline treatment.</summary>
        public ExcelUnderlineStyle? Underline { get; set; }
        /// <summary>Optional font family; null keeps the Normal font.</summary>
        public string? FontName { get; set; }
        /// <summary>Optional positive finite font size in points.</summary>
        public double? FontSize { get; set; }
        /// <summary>Optional font color.</summary>
        public OfficeColor? FontColor { get; set; }
        /// <summary>Optional solid background color.</summary>
        public OfficeColor? BackgroundColor { get; set; }
        /// <summary>Wraps text within the cell.</summary>
        public bool WrapText { get; set; }
        /// <summary>Optional horizontal alignment.</summary>
        public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
        /// <summary>Optional vertical alignment.</summary>
        public ExcelVerticalAlignment? VerticalAlignment { get; set; }
        /// <summary>
        /// Optional Excel number-format code. Null allows tabular writers to select a temporal
        /// format for dates and durations while preserving the other declared properties.
        /// </summary>
        public string? NumberFormat { get; set; }

        internal ExcelStyleDefinition Snapshot() {
            ValidateEnum(Underline, nameof(Underline));
            ValidateEnum(HorizontalAlignment, nameof(HorizontalAlignment));
            ValidateEnum(VerticalAlignment, nameof(VerticalAlignment));
            if (FontSize.HasValue && (FontSize.Value <= 0 || double.IsNaN(FontSize.Value) || double.IsInfinity(FontSize.Value))) {
                throw new ArgumentOutOfRangeException(nameof(FontSize), "Font size must be a positive finite value.");
            }
            ValidateText(FontName, nameof(FontName));
            ValidateText(NumberFormat, nameof(NumberFormat));
            return new ExcelStyleDefinition {
                Bold = Bold, Italic = Italic, Underline = Underline, FontName = FontName?.Trim(),
                FontSize = FontSize, FontColor = FontColor, BackgroundColor = BackgroundColor,
                WrapText = WrapText, HorizontalAlignment = HorizontalAlignment,
                VerticalAlignment = VerticalAlignment, NumberFormat = NumberFormat
            };
        }

        private static void ValidateEnum<T>(T? value, string name) where T : struct {
            if (value.HasValue && !Enum.IsDefined(typeof(T), value.Value)) throw new ArgumentOutOfRangeException(name);
        }

        private static void ValidateText(string? value, string name) {
            if (value == null) return;
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Style text must not be empty.", name);
            try { XmlConvert.VerifyXmlChars(value); }
            catch (XmlException exception) { throw new ArgumentException("Style text must contain valid XML characters.", name, exception); }
        }
    }
}
