using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>A finite, precompiled style catalog shared by all tabular row-writing routes.</summary>
    internal sealed class ExcelTabularStylePlan {
        private readonly Dictionary<string, DeclaredStyle> _styles;
        private ExcelTabularStylePlan(Stylesheet stylesheet, Dictionary<string, DeclaredStyle> styles,
            DeclaredStyle? defaultRowStyle, DeclaredStyle?[] columns) {
            Stylesheet = stylesheet;
            _styles = styles;
            DefaultRowStyle = defaultRowStyle;
            Columns = columns;
        }
        internal Stylesheet Stylesheet { get; }
        internal DeclaredStyle? DefaultRowStyle { get; }
        internal DeclaredStyle?[] Columns { get; }

        internal DeclaredStyle? Resolve(string? name) {
            if (name == null) return null;
            string normalized = ExcelSheet.ValidateNamedStyleName(name);
            if (!_styles.TryGetValue(normalized, out DeclaredStyle? style)) {
                throw new KeyNotFoundException($"Declared style '{normalized}' was not found in Styles.");
            }
            return style;
        }

        internal static ExcelTabularStylePlan? Create(ExcelTabularWriteOptions options, int columnCount, CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            if ((options.Styles == null || options.Styles.Count == 0)
                && options.DefaultRowStyle == null && (options.ColumnStyles == null || options.ColumnStyles.Count == 0)) return null;
            var definitions = new SortedDictionary<string, ExcelStyleDefinition>(StringComparer.OrdinalIgnoreCase);
            if (options.Styles != null) {
                // Each named definition requires at least one XF in addition to the five baseline formats.
                if (options.Styles.Count > 63_995) throw new ArgumentException("Declared styles exceed the 64,000 cell-format export limit.", nameof(options));
                foreach (var item in options.Styles) {
                    ct.ThrowIfCancellationRequested();
                    string name = ExcelSheet.ValidateNamedStyleName(item.Key);
                    if (string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) {
                        throw new ArgumentException("The built-in Normal style cannot be replaced in Styles.", nameof(options));
                    }
                    if (item.Value == null) throw new ArgumentException("Style definitions must not be null.", nameof(options));
                    if (definitions.ContainsKey(name)) throw new ArgumentException("Declared style names must be unique ignoring case and surrounding spaces.", nameof(options));
                    definitions.Add(name, item.Value.Snapshot());
                }
            }
            // Resolve every static selection before generating any package or reading source rows.
            string? rowName = ValidateSelection(options.DefaultRowStyle, definitions);
            var columnNames = new string?[columnCount];
            if (options.ColumnStyles != null) {
                foreach (var item in options.ColumnStyles) {
                    ct.ThrowIfCancellationRequested();
                    if (item.Key < 1 || item.Key > columnCount) throw new ArgumentOutOfRangeException(nameof(options.ColumnStyles), "Style columns must be within the 1-based exported schema.");
                    if (item.Value == null) throw new ArgumentException("Column style names must not be null.", nameof(options));
                    columnNames[item.Key - 1] = ValidateSelection(item.Value, definitions);
                }
            }

            Stylesheet stylesheet = ExcelDocument.CreateDefaultStylesheet();
            uint dateId = ExcelSheet.AddDefinitionNumberFormat(stylesheet, "yyyy-mm-dd hh:mm");
            uint timeId = ExcelSheet.AddDefinitionNumberFormat(stylesheet, "[h]:mm:ss");
            // Preserve the existing temporal indices used by normal row-writer overloads.
            foreach (uint id in new[] { dateId, timeId, 14U, 46U }) AddTemporalFormat(stylesheet, stylesheet.CellFormats!.Elements<CellFormat>().First(), id);
            var styles = new Dictionary<string, DeclaredStyle>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in definitions) {
                ct.ThrowIfCancellationRequested();
                ExcelNamedStyleInfo named = ExcelSheet.DefineDeclaredNamedStyle(stylesheet, item.Key, item.Value, false);
                var format = (CellFormat)stylesheet.CellStyleFormats!.Elements<CellFormat>().ElementAt((int)named.FormatId).CloneNode(true);
                format.FormatId = named.FormatId;
                uint baseId = ExcelSheet.AddDefinitionCellFormat(stylesheet, format);
                string baseAttribute = Attribute(baseId);
                string[] temporal = item.Value.NumberFormat != null
                    ? new[] { baseAttribute, baseAttribute, baseAttribute, baseAttribute }
                    : new[] {
                        Attribute(AddTemporalFormat(stylesheet, format, dateId)),
                        Attribute(AddTemporalFormat(stylesheet, format, timeId)),
                        Attribute(AddTemporalFormat(stylesheet, format, 14U)),
                        Attribute(AddTemporalFormat(stylesheet, format, 46U))
                    };
                styles.Add(item.Key, new DeclaredStyle(baseAttribute, temporal));
                if (stylesheet.CellFormats!.ChildElements.Count > 64_000) {
                    throw new ArgumentException("Declared styles and their temporal variants exceed the 64,000 cell-format export limit.", nameof(options));
                }
            }
            var columns = new DeclaredStyle?[columnCount];
            for (int i = 0; i < columnCount; i++) if (columnNames[i] != null) columns[i] = styles[columnNames[i]!];
            return new ExcelTabularStylePlan(stylesheet, styles, rowName == null ? null : styles[rowName], columns);
        }

        private static string? ValidateSelection(string? name, IDictionary<string, ExcelStyleDefinition> definitions) {
            if (name == null) return null;
            string normalized = ExcelSheet.ValidateNamedStyleName(name);
            if (!definitions.ContainsKey(normalized)) throw new ArgumentException($"Declared style '{normalized}' was not found in Styles.", nameof(name));
            return normalized;
        }

        private static uint AddTemporalFormat(Stylesheet stylesheet, CellFormat source, uint numberFormatId) {
            var format = (CellFormat)source.CloneNode(true);
            format.NumberFormatId = numberFormatId;
            format.ApplyNumberFormat = true;
            return ExcelSheet.AddDefinitionCellFormat(stylesheet, format);
        }

        private static string Attribute(uint index) => " s=\"" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"";

        internal sealed class DeclaredStyle {
            private readonly string[] _temporal;
            internal DeclaredStyle(string attribute, string[] temporal) { Attribute = attribute; _temporal = temporal; }
            internal string Attribute { get; }
            internal string ForValue(string? temporalAttribute) => temporalAttribute switch {
                " s=\"1\"" => _temporal[0], " s=\"2\"" => _temporal[1],
                " s=\"3\"" => _temporal[2], " s=\"4\"" => _temporal[3], _ => Attribute
            };
        }
    }
}
