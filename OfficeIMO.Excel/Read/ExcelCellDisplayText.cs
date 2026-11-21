using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Displayed-text mode for the tabular readers: numeric cells whose style carries a non-General number format
    /// are returned as the text that format displays. Each style resolves to a cached format plan once per workbook;
    /// General and text-formatted cells keep the readers' typed values and their existing conversion route.
    /// </summary>
    internal sealed class ExcelCellDisplayText {
        private readonly ExcelNumberFormatPlan?[] _plans;
        private readonly ExcelDateSystem _dateSystem;

        private ExcelCellDisplayText(ExcelNumberFormatPlan?[] plans, ExcelDateSystem dateSystem, IFormatProvider provider) {
            _plans = plans;
            _dateSystem = dateSystem;
            Provider = provider;
        }

        /// <summary>Culture used for displayed text, including General numbers formatted by the caller.</summary>
        internal IFormatProvider Provider { get; }

        /// <summary>Returns a copy of <paramref name="options"/> that reads cells as displayed text.</summary>
        internal static ExcelReaderOptions Apply(ExcelDocument document, ExcelReaderOptions? options, IFormatProvider provider) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (provider == null) throw new ArgumentNullException(nameof(provider));
            ExcelReaderOptions effective = options?.Clone() ?? new ExcelReaderOptions();
            effective.DisplayText = Create(document.OpenXmlDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet, document.DateSystem, provider);
            return effective;
        }

        internal static ExcelCellDisplayText Create(Stylesheet? stylesheet, ExcelDateSystem dateSystem, IFormatProvider provider) {
            if (stylesheet?.CellFormats == null) return new ExcelCellDisplayText(Array.Empty<ExcelNumberFormatPlan?>(), dateSystem, provider);
            var customCodes = new Dictionary<uint, string>();
            if (stylesheet.NumberingFormats != null) {
                foreach (NumberingFormat numberingFormat in stylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                    if (numberingFormat.NumberFormatId?.Value is uint id && numberingFormat.FormatCode?.Value is string code) {
                        customCodes[id] = code;
                    }
                }
            }

            // Styles that share a number format share one plan.
            var plansByFormat = new Dictionary<uint, ExcelNumberFormatPlan?>();
            var plans = new List<ExcelNumberFormatPlan?>();
            foreach (CellFormat cellFormat in stylesheet.CellFormats.Elements<CellFormat>()) {
                uint numberFormatId = cellFormat.NumberFormatId?.Value ?? 0U;
                if (!plansByFormat.TryGetValue(numberFormatId, out ExcelNumberFormatPlan? plan)) {
                    string? code = customCodes.TryGetValue(numberFormatId, out string? custom)
                        ? custom
                        : ExcelBuiltInNumberFormats.GetCode(numberFormatId);
                    plan = ExcelNumberFormatPlan.Create(numberFormatId, code, dateSystem, provider);
                    plansByFormat[numberFormatId] = plan;
                }
                plans.Add(plan);
            }
            return new ExcelCellDisplayText(plans.ToArray(), dateSystem, provider);
        }

        /// <summary>True when numeric cells with this style display differently from their General text.</summary>
        internal bool IsFormatted(uint styleIndex) => styleIndex < (uint)_plans.Length && _plans[styleIndex] != null;

        internal bool IsFormatted(uint? styleIndex) => styleIndex.HasValue && IsFormatted(styleIndex.Value);

        /// <summary>Formats a numeric cell value; the style must satisfy <see cref="IsFormatted(uint)"/>.</summary>
        internal string FormatNumber(double value, uint styleIndex) => _plans[styleIndex]!.Format(value);

        /// <summary>Formats an ISO 8601 date cell (<c>t="d"</c>) through its style, when the style is formatted.</summary>
        internal bool TryFormatIsoDate(string? raw, uint? styleIndex, out string text) {
            text = string.Empty;
            if (raw == null || !IsFormatted(styleIndex)
                || !ExcelIsoDateText.TryParseForDisplay(raw, _dateSystem, out DateTime date, out double serial)) {
                return false;
            }
            ExcelNumberFormatPlan plan = _plans[styleIndex!.Value]!;
            string? formatted = ExcelNumberFormatDisplay.FormatDateTimeText(date, plan.NumberFormatId, plan.FormatCode,
                serial, Provider);
            if (formatted == null) return false;
            text = formatted;
            return true;
        }

        /// <summary>
        /// Displays a typed value that no number format applies to: numbers as General with up to 15 significant digits,
        /// booleans as TRUE or FALSE. Returns null for other values, whose own text is already what Excel displays.
        /// </summary>
        internal string? FormatUnstyled(object? value) => value switch {
            double number => OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.FormatGeneralValue(number, Provider),
            decimal amount => OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.FormatGeneralValue((double)amount, Provider),
            bool flag => flag ? "TRUE" : "FALSE",
            _ => null
        };

        /// <summary>Formats a raw numeric or ISO date cell value when its style is formatted.</summary>
        internal bool TryFormatRaw(bool numericCell, bool isoDateCell, string? raw, uint? styleIndex, out string text) {
            text = string.Empty;
            if (raw == null || !IsFormatted(styleIndex)) return false;
            if (isoDateCell) return TryFormatIsoDate(raw, styleIndex, out text);
            if (!numericCell || !double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double number)) {
                return false;
            }
            text = FormatNumber(number, styleIndex!.Value);
            return true;
        }
    }
}
