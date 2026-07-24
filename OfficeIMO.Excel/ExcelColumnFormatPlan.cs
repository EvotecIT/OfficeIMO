using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes reusable header-based column number formats for worksheet exports.
    /// </summary>
    public sealed class ExcelColumnFormatPlan {
        private readonly List<ExcelColumnFormatRule> _rules = new();

        /// <summary>
        /// Gets the configured column format rules.
        /// </summary>
        public IReadOnlyList<ExcelColumnFormatRule> Rules => _rules;

        /// <summary>
        /// Gets the number of configured rules.
        /// </summary>
        public int Count => _rules.Count;

        /// <summary>
        /// Adds a preset number format for a column resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Add(
            string header,
            ExcelNumberPreset preset,
            int decimals = 2,
            CultureInfo? culture = null,
            bool includeHeader = false,
            bool autoFit = false) {
            _rules.Add(ExcelColumnFormatRule.FromPreset(header, preset, decimals, culture, includeHeader, autoFit));
            return this;
        }

        /// <summary>
        /// Adds a custom Excel number format for a column resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan AddFormat(string header, string numberFormat, bool includeHeader = false, bool autoFit = false) {
            _rules.Add(ExcelColumnFormatRule.FromFormat(header, numberFormat, includeHeader, autoFit));
            return this;
        }

        /// <summary>
        /// Adds text formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Text(params string[] headers) => AddPreset(headers, ExcelNumberPreset.Text, 0);

        /// <summary>
        /// Adds decimal number formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Number(params string[] headers) => AddPreset(headers, ExcelNumberPreset.Decimal, 2);

        /// <summary>
        /// Adds decimal number formatting with explicit precision for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Number(int decimals, params string[] headers) => AddPreset(headers, ExcelNumberPreset.Decimal, decimals);

        /// <summary>
        /// Adds integer formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Integer(params string[] headers) => AddPreset(headers, ExcelNumberPreset.Integer, 0);

        /// <summary>
        /// Adds percent formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Percent(params string[] headers) => AddPreset(headers, ExcelNumberPreset.Percent, 0);

        /// <summary>
        /// Adds percent formatting with explicit precision for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Percent(int decimals, params string[] headers) => AddPreset(headers, ExcelNumberPreset.Percent, decimals);

        /// <summary>
        /// Adds currency formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Currency(params string[] headers) => Currency(2, null, headers);

        /// <summary>
        /// Adds culture-aware currency formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Currency(CultureInfo? culture, params string[] headers) => Currency(2, culture, headers);

        /// <summary>
        /// Adds currency formatting with explicit precision for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Currency(int decimals, params string[] headers) => Currency(decimals, null, headers);

        /// <summary>
        /// Adds culture-aware currency formatting with explicit precision for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Currency(int decimals, CultureInfo? culture, params string[] headers) {
            foreach (string header in NormalizeHeaders(headers)) {
                Add(header, ExcelNumberPreset.Currency, decimals, culture);
            }
            return this;
        }

        /// <summary>
        /// Adds date formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan Date(params string[] headers) => AddPreset(headers, ExcelNumberPreset.DateShort, 0);

        /// <summary>
        /// Adds date and time formatting for columns resolved by header text.
        /// </summary>
        public ExcelColumnFormatPlan DateTime(params string[] headers) => AddPreset(headers, ExcelNumberPreset.DateTime, 0);

        private ExcelColumnFormatPlan AddPreset(IEnumerable<string> headers, ExcelNumberPreset preset, int decimals) {
            foreach (string header in NormalizeHeaders(headers)) {
                Add(header, preset, decimals);
            }
            return this;
        }

        private static IEnumerable<string> NormalizeHeaders(IEnumerable<string>? headers) {
            if (headers == null) {
                yield break;
            }

            foreach (string? header in headers) {
                if (!string.IsNullOrWhiteSpace(header)) {
                    yield return header;
                }
            }
        }
    }

    /// <summary>
    /// Describes one header-based column number format.
    /// </summary>
    public sealed class ExcelColumnFormatRule {
        private ExcelColumnFormatRule(
            string header,
            ExcelNumberPreset? preset,
            string? numberFormat,
            int decimals,
            CultureInfo? culture,
            bool includeHeader,
            bool autoFit) {
            if (string.IsNullOrWhiteSpace(header)) {
                throw new ArgumentException("Header cannot be empty.", nameof(header));
            }

            if (preset == null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("A preset or custom number format is required.", nameof(numberFormat));
            }

            if (decimals < 0 || decimals > ExcelNumberFormats.MaximumDecimalPlaces) {
                throw new ArgumentOutOfRangeException(nameof(decimals), $"Decimals must be between 0 and {ExcelNumberFormats.MaximumDecimalPlaces}.");
            }

            Header = header;
            Preset = preset;
            NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat;
            Decimals = decimals;
            Culture = culture;
            IncludeHeader = includeHeader;
            AutoFit = autoFit;
        }

        /// <summary>Header caption used to resolve the column.</summary>
        public string Header { get; }

        /// <summary>Preset number format, when the rule was created from a preset.</summary>
        public ExcelNumberPreset? Preset { get; }

        /// <summary>Custom number format, when supplied directly.</summary>
        public string? NumberFormat { get; }

        /// <summary>Decimal places used by preset formats that support precision.</summary>
        public int Decimals { get; }

        /// <summary>Culture used for culture-aware preset formats such as currency.</summary>
        public CultureInfo? Culture { get; }

        /// <summary>Whether the header cell should also receive the format.</summary>
        public bool IncludeHeader { get; }

        /// <summary>Whether the resolved column should be auto-fit after formatting.</summary>
        public bool AutoFit { get; }

        /// <summary>
        /// Creates a preset-based column format rule.
        /// </summary>
        public static ExcelColumnFormatRule FromPreset(
            string header,
            ExcelNumberPreset preset,
            int decimals = 2,
            CultureInfo? culture = null,
            bool includeHeader = false,
            bool autoFit = false) {
            return new ExcelColumnFormatRule(header, preset, null, decimals, culture, includeHeader, autoFit);
        }

        /// <summary>
        /// Creates a custom number-format rule.
        /// </summary>
        public static ExcelColumnFormatRule FromFormat(string header, string numberFormat, bool includeHeader = false, bool autoFit = false) {
            return new ExcelColumnFormatRule(header, null, numberFormat, 0, null, includeHeader, autoFit);
        }

        /// <summary>
        /// Resolves the Excel number format code for this rule.
        /// </summary>
        public string ResolveNumberFormat() {
            return NumberFormat ?? ExcelNumberFormats.Get(Preset ?? ExcelNumberPreset.General, Decimals, Culture);
        }
    }

    /// <summary>
    /// Result for one applied or skipped column format rule.
    /// </summary>
    public sealed class ExcelColumnFormatResult {
        internal ExcelColumnFormatResult(string header, int? columnIndex, bool applied, string numberFormat, string? warning) {
            Header = header;
            ColumnIndex = columnIndex;
            Applied = applied;
            NumberFormat = numberFormat;
            Warning = warning;
        }

        /// <summary>Header requested by the format rule.</summary>
        public string Header { get; }

        /// <summary>One-based column index when the header was resolved.</summary>
        public int? ColumnIndex { get; }

        /// <summary>Whether the rule was applied.</summary>
        public bool Applied { get; }

        /// <summary>Number format code resolved for the rule.</summary>
        public string NumberFormat { get; }

        /// <summary>Warning for skipped rules, when any.</summary>
        public string? Warning { get; }
    }
}
