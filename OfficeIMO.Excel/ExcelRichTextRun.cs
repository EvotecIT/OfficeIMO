using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a run of rich text inside a cell.
    /// </summary>
    public sealed class ExcelRichTextRun {
        private bool _bold;
        private bool _italic;
        private bool _underline;
        private bool _strikethrough;
        private ExcelUnderlineStyle? _underlineStyle;
        private OfficeLength? _fontSize;

        /// <summary>
        /// Creates a rich text run with the supplied text.
        /// </summary>
        public ExcelRichTextRun(string text) {
            Text = text ?? string.Empty;
        }

        /// <summary>Gets or sets the run text.</summary>
        public string Text { get; set; }
        /// <summary>Gets or sets whether the run is bold.</summary>
        public bool Bold {
            get => _bold;
            set {
                _bold = value;
                BoldSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is italic.</summary>
        public bool Italic {
            get => _italic;
            set {
                _italic = value;
                ItalicSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is underlined.</summary>
        public bool Underline {
            get => _underline;
            set {
                _underline = value;
                if (!value) {
                    _underlineStyle = ExcelUnderlineStyle.None;
                } else if (!_underlineStyle.HasValue || _underlineStyle == ExcelUnderlineStyle.None) {
                    _underlineStyle = ExcelUnderlineStyle.Single;
                }
                UnderlineSpecified = true;
            }
        }
        /// <summary>Gets or sets whether the run is struck through.</summary>
        public bool Strikethrough {
            get => _strikethrough;
            set {
                _strikethrough = value;
                StrikethroughSpecified = true;
            }
        }
        /// <summary>Gets or sets the run underline style.</summary>
        public ExcelUnderlineStyle? UnderlineStyle {
            get => _underlineStyle;
            set {
                _underlineStyle = value;
                if (value.HasValue) {
                    _underline = value != ExcelUnderlineStyle.None;
                    UnderlineSpecified = true;
                }
            }
        }
        /// <summary>Gets or sets the literal or theme-based run text color.</summary>
        public OfficeColor? Color { get; set; }
        /// <summary>Gets or sets the run typeface family name.</summary>
        public string? FontFamily { get; set; }
        /// <summary>Gets or sets the positive physical run font size; null inherits the cell font size.</summary>
        public OfficeLength? FontSize {
            get => _fontSize;
            set {
                if (value.HasValue && value.Value <= OfficeLength.Zero)
                    throw new ArgumentOutOfRangeException(nameof(FontSize), "Font size must be positive.");
                _fontSize = value;
            }
        }
        /// <summary>Gets or sets the run vertical text alignment.</summary>
        public ExcelVerticalTextAlignment? VerticalTextAlignment { get; set; }
        /// <summary>Gets or sets whether the run font uses outline text.</summary>
        public bool Outline { get; set; }
        /// <summary>Gets or sets whether the run font uses shadow text.</summary>
        public bool Shadow { get; set; }
        /// <summary>Gets or sets whether the run font uses condensed text.</summary>
        public bool Condense { get; set; }
        /// <summary>Gets or sets whether the run font uses extended text.</summary>
        public bool Extend { get; set; }
        /// <summary>Gets or sets the run font family classification byte.</summary>
        public byte? FontFamilyClassification { get; set; }
        /// <summary>Gets or sets the run font character set byte.</summary>
        public byte? FontCharacterSet { get; set; }

        internal bool BoldSpecified { get; private set; }
        internal bool ItalicSpecified { get; private set; }
        internal bool UnderlineSpecified { get; private set; }
        internal bool StrikethroughSpecified { get; private set; }

        /// <summary>
        /// Creates a plain rich text run.
        /// </summary>
        public static ExcelRichTextRun Plain(string text) => new ExcelRichTextRun(text);

        /// <summary>
        /// Changes the stored run text casing while preserving rich-text formatting.
        /// </summary>
        public ExcelRichTextRun TransformTextCase(OfficeTextCase textCase, CultureInfo? culture = null) {
            Text = OfficeTextCaseTransformer.Apply(Text, textCase, culture);
            return this;
        }

        /// <summary>Projects native run properties without losing explicit disabled values.</summary>
        internal static ExcelRichTextRun FromOpenXml(string text, RunProperties? properties, string? resolvedFontColor = null) {
            var result = new ExcelRichTextRun(text) {
                Color = ExcelColor.FromSpreadsheetColor(properties?.GetFirstChild<Color>(), resolvedFontColor),
                FontFamily = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value is double points ? OfficeLength.Points(points) : (OfficeLength?)null,
                VerticalTextAlignment = GetVerticalTextAlignment(properties),
                Outline = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Outline>()),
                Shadow = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Shadow>()),
                Condense = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Condense>()),
                Extend = ExcelOpenXmlFontProperty.IsEnabled(properties?.GetFirstChild<Extend>()),
                FontFamilyClassification = GetFontFamily(properties),
                FontCharacterSet = GetFontCharacterSet(properties)
            };

            DocumentFormat.OpenXml.Spreadsheet.Bold? bold = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Bold>();
            if (bold != null) result.Bold = ExcelOpenXmlFontProperty.IsEnabled(bold);
            DocumentFormat.OpenXml.Spreadsheet.Italic? italic = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Italic>();
            if (italic != null) result.Italic = ExcelOpenXmlFontProperty.IsEnabled(italic);
            DocumentFormat.OpenXml.Spreadsheet.Underline? underline = properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>();
            if (underline != null) {
                result.Underline = ExcelOpenXmlFontProperty.IsUnderlineEnabled(underline);
                result.UnderlineStyle = GetUnderlineStyle(properties);
            }
            Strike? strike = properties?.GetFirstChild<Strike>();
            if (strike != null) result.Strikethrough = ExcelOpenXmlFontProperty.IsEnabled(strike);
            return result;
        }

        /// <summary>Copies a run while retaining which direct Boolean properties were present.</summary>
        internal ExcelRichTextRun Clone(string? text = null) {
            var result = new ExcelRichTextRun(text ?? Text) {
                Color = Color,
                FontFamily = FontFamily,
                FontSize = FontSize,
                VerticalTextAlignment = VerticalTextAlignment,
                Outline = Outline,
                Shadow = Shadow,
                Condense = Condense,
                Extend = Extend,
                FontFamilyClassification = FontFamilyClassification,
                FontCharacterSet = FontCharacterSet
            };
            if (BoldSpecified) result.Bold = Bold;
            if (ItalicSpecified) result.Italic = Italic;
            if (UnderlineSpecified) {
                result.Underline = Underline;
                result.UnderlineStyle = UnderlineStyle;
            }
            if (StrikethroughSpecified) result.Strikethrough = Strikethrough;
            return result;
        }

        internal static void AppendFontMetadata(RunProperties properties, ExcelRichTextRun run) {
            if (run.VerticalTextAlignment.HasValue) {
                properties.Append(new VerticalTextAlignment { Val = run.VerticalTextAlignment.Value.ToOpenXml() });
            }

            if (run.Outline) {
                properties.Append(new Outline());
            }

            if (run.Shadow) {
                properties.Append(new Shadow());
            }

            if (run.Condense) {
                properties.Append(new Condense());
            }

            if (run.Extend) {
                properties.Append(new Extend());
            }

            if (run.FontFamilyClassification.HasValue) {
                properties.Append(new DocumentFormat.OpenXml.Spreadsheet.FontFamily { Val = run.FontFamilyClassification.Value });
            }

            if (run.FontCharacterSet.HasValue) {
                var charset = new OpenXmlUnknownElement(string.Empty, "charset", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                charset.SetAttribute(new OpenXmlAttribute("val", string.Empty, run.FontCharacterSet.Value.ToString(CultureInfo.InvariantCulture)));
                properties.Append(charset);
            }
        }

        internal static byte? GetFontFamily(RunProperties? properties) {
            return TryGetByte(properties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontFamily>()?.Val?.Value, out byte family)
                ? family
                : null;
        }

        internal static byte? GetFontCharacterSet(RunProperties? properties) {
            OpenXmlElement? charset = properties?.ChildElements.FirstOrDefault(child =>
                string.Equals(child.LocalName, "charset", StringComparison.Ordinal)
                && string.Equals(child.NamespaceUri, "http://schemas.openxmlformats.org/spreadsheetml/2006/main", StringComparison.Ordinal));
            return TryGetByte(charset?.GetAttribute("val", string.Empty).Value, out byte characterSet)
                ? characterSet
                : null;
        }

        internal static ExcelVerticalTextAlignment? GetVerticalTextAlignment(RunProperties? properties) {
            return properties?.GetFirstChild<VerticalTextAlignment>()?.Val?.Value.ToOfficeEnum();
        }

        internal static ExcelUnderlineStyle? GetUnderlineStyle(RunProperties? properties) {
            Underline? underline = properties?.GetFirstChild<Underline>();
            if (underline == null) {
                return null;
            }

            return underline.Val?.Value.ToOfficeEnum() ?? ExcelUnderlineStyle.Single;
        }

        private static bool TryGetByte(uint? value, out byte result) {
            if (value.HasValue && value.Value <= byte.MaxValue) {
                result = checked((byte)value.Value);
                return true;
            }

            result = 0;
            return false;
        }

        private static bool TryGetByte(int? value, out byte result) {
            if (value.HasValue && value.Value >= 0 && value.Value <= byte.MaxValue) {
                result = checked((byte)value.Value);
                return true;
            }

            result = 0;
            return false;
        }

        private static bool TryGetByte(string? value, out byte result) {
            if (!string.IsNullOrWhiteSpace(value)
                && uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint parsed)
                && parsed <= byte.MaxValue) {
                result = checked((byte)parsed);
                return true;
            }

            result = 0;
            return false;
        }
    }
}
