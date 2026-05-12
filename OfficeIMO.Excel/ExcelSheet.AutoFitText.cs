using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool IsSimpleAutoFitCharacter(char value)
            => (value >= '0' && value <= '9')
            || value == '.'
            || value == ','
            || value == '-'
            || value == '+'
            || value == '/'
            || value == ':'
            || value == ' '
            || value == '%';

        private AutoFitTextContext CreateAutoFitTextContext() {
            var stylesheet = WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet;
            return new AutoFitTextContext(
                _excelDocument.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList(),
                stylesheet?.CellFormats?.Elements<CellFormat>().ToList(),
                stylesheet?.NumberingFormats?.Elements<NumberingFormat>().ToDictionary(
                    nf => nf.NumberFormatId?.Value ?? 0U,
                    nf => nf.FormatCode?.Value));
        }

        private string GetCellAutoFitText(Cell cell, AutoFitTextContext? context = null) {
            var dataType = cell.DataType?.Value;

            if (dataType == XCellValues.SharedString) {
                return GetCachedSharedStringText(cell, context);
            }

            if (dataType == XCellValues.InlineString) {
                return GetInlineStringText(cell.InlineString);
            }

            string raw = cell.CellValue?.InnerText ?? string.Empty;
            if (string.IsNullOrEmpty(raw)) {
                return string.Empty;
            }

            if (dataType == XCellValues.Boolean) {
                return raw == "1" ? "TRUE" : raw == "0" ? "FALSE" : raw;
            }

            if (dataType == XCellValues.Error) {
                return raw;
            }

            if (dataType == null || dataType == XCellValues.Number) {
                return FormatAutoFitNumericText(cell, raw, context);
            }

            return raw;
        }

        private string FormatAutoFitNumericText(Cell cell, string raw, AutoFitTextContext? context) {
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                return raw;
            }

            uint numberFormatId = GetCellNumberFormatId(cell, context);
            string? formatCode = GetNumberFormatCode(numberFormatId, context);

            if (numberFormatId == 0U || string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                return raw;
            }

            if (IsDateNumberFormat(numberFormatId, formatCode)) {
                return FormatAutoFitDateValue(value, numberFormatId, formatCode!);
            }

            return FormatAutoFitNumberValue(value, numberFormatId, formatCode!) ?? raw;
        }

        private uint GetCellNumberFormatId(Cell cell, AutoFitTextContext? context) {
            if (cell.StyleIndex == null) {
                return 0U;
            }

            uint styleIndex = cell.StyleIndex.Value;
            if (context != null && context.NumberFormatIdsByStyle.TryGetValue(styleIndex, out uint cachedFormatId)) {
                return cachedFormatId;
            }

            CellFormat? cellFormat = null;
            if (context?.CellFormats != null) {
                cellFormat = styleIndex < context.CellFormats.Count ? context.CellFormats[(int)styleIndex] : null;
            } else {
                var cellFormats = WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet?.CellFormats;
                cellFormat = cellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
            }

            uint numberFormatId = cellFormat?.NumberFormatId?.Value ?? 0U;
            if (context != null) {
                context.NumberFormatIdsByStyle[styleIndex] = numberFormatId;
            }

            return numberFormatId;
        }

        private string? GetNumberFormatCode(uint numberFormatId, AutoFitTextContext? context) {
            if (context != null && context.NumberFormatCodes.TryGetValue(numberFormatId, out string? cachedCode)) {
                return cachedCode;
            }

            string? code = GetNumberFormatCodeCore(numberFormatId, context);
            if (context != null) {
                context.NumberFormatCodes[numberFormatId] = code;
            }

            return code;
        }

        private string? GetNumberFormatCodeCore(uint numberFormatId, AutoFitTextContext? context) {
            string? builtIn = GetBuiltInNumberFormatCode(numberFormatId);
            if (builtIn != null) {
                return builtIn;
            }

            if (context?.CustomNumberFormats != null) {
                return context.CustomNumberFormats.TryGetValue(numberFormatId, out string? formatCode) ? formatCode : null;
            }

            var numberingFormats = WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet?.NumberingFormats;
            if (numberingFormats == null) {
                return null;
            }

            foreach (var numberingFormat in numberingFormats.Elements<NumberingFormat>()) {
                if (numberingFormat.NumberFormatId?.Value == numberFormatId) {
                    return numberingFormat.FormatCode?.Value;
                }
            }

            return null;
        }

        private static string? GetBuiltInNumberFormatCode(uint id) {
            switch (id) {
                case 0: return "General";
                case 1: return "0";
                case 2: return "0.00";
                case 3: return "#,##0";
                case 4: return "#,##0.00";
                case 9: return "0%";
                case 10: return "0.00%";
                case 11: return "0.00E+00";
                case 14: return "m/d/yyyy";
                case 15: return "d-mmm-yy";
                case 16: return "d-mmm";
                case 17: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: return "h:mm";
                case 21: return "h:mm:ss";
                case 22: return "m/d/yyyy h:mm";
                case 37: return "#,##0;(#,##0)";
                case 38: return "#,##0;[Red](#,##0)";
                case 39: return "#,##0.00;(#,##0.00)";
                case 40: return "#,##0.00;[Red](#,##0.00)";
                case 45: return "mm:ss";
                case 46: return "[h]:mm:ss";
                case 47: return "mm:ss.0";
                case 49: return "@";
                default: return null;
            }
        }

        private static bool IsDateNumberFormat(uint numberFormatId, string? formatCode)
            => numberFormatId is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22 or 45 or 46 or 47
            || ExcelNumberFormatClassifier.LooksLikeDateFormat(formatCode);

        private static string FormatAutoFitDateValue(double value, uint numberFormatId, string formatCode) {
            if (numberFormatId == 46U || formatCode.IndexOf("[h]", StringComparison.OrdinalIgnoreCase) >= 0) {
                TimeSpan duration = TimeSpan.FromDays(value);
                int totalHours = (int)Math.Floor(duration.TotalHours);
                return string.Format(CultureInfo.InvariantCulture, "{0}:{1:00}:{2:00}", totalHours, Math.Abs(duration.Minutes), Math.Abs(duration.Seconds));
            }

            DateTime date;
            try {
                date = DateTime.FromOADate(value);
            } catch {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            switch (numberFormatId) {
                case 14: return date.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                case 15: return date.ToString("d-MMM-yy", CultureInfo.InvariantCulture);
                case 16: return date.ToString("d-MMM", CultureInfo.InvariantCulture);
                case 17: return date.ToString("MMM-yy", CultureInfo.InvariantCulture);
                case 18: return date.ToString("h:mm tt", CultureInfo.InvariantCulture);
                case 19: return date.ToString("h:mm:ss tt", CultureInfo.InvariantCulture);
                case 20: return date.ToString("H:mm", CultureInfo.InvariantCulture);
                case 21: return date.ToString("H:mm:ss", CultureInfo.InvariantCulture);
                case 22: return date.ToString("M/d/yyyy H:mm", CultureInfo.InvariantCulture);
                case 45: return date.ToString("mm:ss", CultureInfo.InvariantCulture);
                case 47: return date.ToString("mm:ss.0", CultureInfo.InvariantCulture);
                default:
                    return date.ToString(TranslateExcelDateFormat(formatCode), CultureInfo.InvariantCulture);
            }
        }

        private static string TranslateExcelDateFormat(string formatCode) {
            string section = SelectNumberFormatSection(formatCode, 0);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm:ss")) return "yyyy-MM-dd HH:mm:ss";
            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm")) return "yyyy-MM-dd HH:mm";
            if (lower.Contains("yyyy-mm-dd")) return "yyyy-MM-dd";
            if (lower.Contains("dd/mm/yyyy")) return "dd/MM/yyyy";
            if (lower.Contains("mm/dd/yyyy")) return "MM/dd/yyyy";
            if (lower.Contains("m/d/yyyy")) return "M/d/yyyy";
            if (lower.Contains("d-mmm-yy")) return "d-MMM-yy";
            if (lower.Contains("mmm-yy")) return "MMM-yy";
            if (lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "h:mm:ss tt";
            if (lower.Contains("h:mm") && lower.Contains("am/pm")) return "h:mm tt";
            if (lower.Contains("hh:mm:ss")) return "HH:mm:ss";
            if (lower.Contains("h:mm:ss")) return "H:mm:ss";
            if (lower.Contains("hh:mm")) return "HH:mm";
            if (lower.Contains("h:mm")) return "H:mm";
            return "M/d/yyyy";
        }

        private static string? FormatAutoFitNumberValue(double value, uint numberFormatId, string formatCode) {
            string section = SelectNumberFormatSection(formatCode, value < 0 ? 1 : value == 0 ? 2 : 0);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (numberFormatId == 49U || lower.Contains("@")) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            if (lower.Contains("e+")) {
                int decimals = CountDecimalPlaces(lower);
                return value.ToString("E" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            }

            bool percent = lower.Contains("%");
            bool thousands = lower.Contains("#,##") || lower.Contains(",##");
            bool currency = normalized.IndexOf('$') >= 0 || normalized.IndexOf('€') >= 0 || normalized.IndexOf('£') >= 0;
            int decimalPlaces = CountDecimalPlaces(lower);
            double displayValue = percent ? value * 100.0 : value;
            string numericFormat = thousands || currency
                ? "N" + decimalPlaces.ToString(CultureInfo.InvariantCulture)
                : "F" + decimalPlaces.ToString(CultureInfo.InvariantCulture);
            string text = displayValue.ToString(numericFormat, CultureInfo.InvariantCulture);

            if (currency) {
                char symbol = normalized.IndexOf('€') >= 0 ? '€' : normalized.IndexOf('£') >= 0 ? '£' : '$';
                text = symbol + text;
            }

            if (percent) {
                text += "%";
            }

            if (value < 0 && normalized.Contains("(")) {
                text = "(" + text.TrimStart('-') + ")";
            }

            return text;
        }

        private static int CountDecimalPlaces(string formatCode) {
            int dot = formatCode.IndexOf('.');
            if (dot < 0) {
                return 0;
            }

            int count = 0;
            for (int i = dot + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0' || ch == '#') {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string SelectNumberFormatSection(string formatCode, int preferredSection) {
            string[] sections = formatCode.Split(';');
            if (sections.Length == 0) {
                return formatCode;
            }

            if (preferredSection >= 0 && preferredSection < sections.Length && !string.IsNullOrWhiteSpace(sections[preferredSection])) {
                return sections[preferredSection];
            }

            return sections[0];
        }

        private static string StripNumberFormatDecorations(string formatCode) {
            var builder = new StringBuilder(formatCode.Length);
            bool inQuote = false;

            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote && ch == '[') {
                    int close = formatCode.IndexOf(']', i + 1);
                    if (close >= 0) {
                        string token = formatCode.Substring(i + 1, close - i - 1);
                        if (token.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S')) {
                            builder.Append('[').Append(token).Append(']');
                        }

                        i = close;
                        continue;
                    }
                }

                if (!inQuote && (ch == '\\' || ch == '_' || ch == '*')) {
                    if (i + 1 < formatCode.Length) {
                        i++;
                    }
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static bool TryGetSharedStringIndex(Cell cell, out int id) {
            var raw = cell.CellValue?.InnerText;
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out id);
        }

        private string GetCachedSharedStringText(Cell cell, AutoFitTextContext? context) {
            if (!TryGetSharedStringIndex(cell, out int id)) {
                return string.Empty;
            }

            if (context != null && context.SharedStringTexts.TryGetValue(id, out string? cachedText)) {
                return cachedText;
            }

            var sharedStrings = context?.SharedStrings ?? _excelDocument.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList();
            string text = sharedStrings != null && id >= 0 && id < sharedStrings.Count
                ? GetSharedStringText(sharedStrings[id])
                : string.Empty;

            if (context != null) {
                context.SharedStringTexts[id] = text;
            }

            return text;
        }

        private IReadOnlyList<AutoFitTextRun>? GetCellAutoFitRichTextRuns(Cell cell, AutoFitTextContext? context = null) {
            if (cell.DataType?.Value == XCellValues.SharedString) {
                if (!TryGetSharedStringIndex(cell, out int id)) {
                    return null;
                }

                if (context != null && context.SharedStringRichTextRuns.TryGetValue(id, out var cachedRuns)) {
                    return cachedRuns;
                }

                var sharedStrings = context?.SharedStrings ?? _excelDocument.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList();
                var runs = sharedStrings != null && id >= 0 && id < sharedStrings.Count
                    ? GetSharedStringRichTextRuns(sharedStrings[id])
                    : null;

                if (context != null) {
                    context.SharedStringRichTextRuns[id] = runs;
                }

                return runs;
            }

            if (cell.DataType?.Value == XCellValues.InlineString) {
                return GetInlineStringRichTextRuns(cell.InlineString);
            }

            return null;
        }

        private sealed class AutoFitTextContext {
            internal AutoFitTextContext(
                IReadOnlyList<SharedStringItem>? sharedStrings,
                IReadOnlyList<CellFormat>? cellFormats,
                Dictionary<uint, string?>? customNumberFormats) {
                SharedStrings = sharedStrings;
                CellFormats = cellFormats;
                CustomNumberFormats = customNumberFormats;
            }

            internal IReadOnlyList<SharedStringItem>? SharedStrings { get; }
            internal IReadOnlyList<CellFormat>? CellFormats { get; }
            internal Dictionary<uint, string?>? CustomNumberFormats { get; }
            internal Dictionary<uint, uint> NumberFormatIdsByStyle { get; } = new Dictionary<uint, uint>();
            internal Dictionary<uint, string?> NumberFormatCodes { get; } = new Dictionary<uint, string?>();
            internal Dictionary<int, string> SharedStringTexts { get; } = new Dictionary<int, string>();
            internal Dictionary<int, IReadOnlyList<AutoFitTextRun>?> SharedStringRichTextRuns { get; } = new Dictionary<int, IReadOnlyList<AutoFitTextRun>?>();
        }

        private static IReadOnlyList<AutoFitTextRun>? GetSharedStringRichTextRuns(SharedStringItem item) {
            if (item.Text != null) {
                return null;
            }

            var runs = item.Elements<Run>().ToList();
            if (runs.Count == 0) {
                return null;
            }

            return CreateAutoFitTextRuns(runs);
        }

        private static IReadOnlyList<AutoFitTextRun>? GetInlineStringRichTextRuns(InlineString? inlineString) {
            if (inlineString == null) {
                return null;
            }

            if (inlineString.Text != null) {
                return null;
            }

            var runs = inlineString.Elements<Run>().ToList();
            if (runs.Count == 0) {
                return null;
            }

            return CreateAutoFitTextRuns(runs);
        }

        private static IReadOnlyList<AutoFitTextRun> CreateAutoFitTextRuns(IReadOnlyList<Run> runs) {
            var result = new List<AutoFitTextRun>(runs.Count);
            foreach (var run in runs) {
                string text = run.Text?.Text ?? string.Empty;
                if (text.Length == 0) {
                    continue;
                }

                result.Add(AutoFitTextRun.Create(text, run.RunProperties));
            }

            return result;
        }

        private static string GetSharedStringText(SharedStringItem item) {
            if (item.Text != null) {
                return item.Text.Text ?? string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var text in item.Descendants<Text>()) {
                sb.Append(text.Text);
            }

            return sb.ToString();
        }

        private static string GetInlineStringText(InlineString? inlineString) {
            if (inlineString == null) {
                return string.Empty;
            }

            if (inlineString.Text != null) {
                return inlineString.Text.Text ?? string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var run in inlineString.Elements<Run>()) {
                if (run.Text != null) {
                    sb.Append(run.Text.Text);
                }
            }

            return sb.ToString();
        }

        private readonly struct AutoFitTextRun {
            private AutoFitTextRun(string text, string? familyName, double? size, bool? bold, bool? italic, bool? underline) {
                Text = text;
                FamilyName = familyName;
                Size = size;
                Bold = bold;
                Italic = italic;
                Underline = underline;
            }

            internal string Text { get; }
            private string? FamilyName { get; }
            private double? Size { get; }
            private bool? Bold { get; }
            private bool? Italic { get; }
            private bool? Underline { get; }

            internal string Signature => string.Join("\u001f", new[] {
                Text,
                FamilyName ?? string.Empty,
                Size?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                Bold?.ToString() ?? string.Empty,
                Italic?.ToString() ?? string.Empty,
                Underline?.ToString() ?? string.Empty
            });

            internal static AutoFitTextRun Create(string text, RunProperties? properties) {
                if (properties == null) {
                    return new AutoFitTextRun(text, null, null, null, null, null);
                }

                string? familyName = properties.GetFirstChild<RunFont>()?.Val?.Value;
                double? size = properties.GetFirstChild<FontSize>()?.Val?.Value;
                bool? bold = GetOptionalBoolean(properties.GetFirstChild<Bold>());
                bool? italic = GetOptionalBoolean(properties.GetFirstChild<Italic>());
                bool? underline = GetOptionalBoolean(properties.GetFirstChild<Underline>());
                return new AutoFitTextRun(text, familyName, size, bold, italic, underline);
            }

            internal OfficeFontInfo CreateFontInfo(OfficeFontInfo fallback) {
                var style = fallback.Style;
                style = ApplyStyle(style, OfficeFontStyle.Bold, Bold);
                style = ApplyStyle(style, OfficeFontStyle.Italic, Italic);
                style = ApplyStyle(style, OfficeFontStyle.Underline, Underline);

                return new OfficeFontInfo(
                    string.IsNullOrWhiteSpace(FamilyName) ? fallback.FamilyName : FamilyName,
                    Size.HasValue && Size.Value > 0 ? Size.Value : fallback.Size,
                    style);
            }

            private static OfficeFontStyle ApplyStyle(OfficeFontStyle style, OfficeFontStyle flag, bool? value) {
                if (!value.HasValue) {
                    return style;
                }

                return value.Value ? style | flag : style & ~flag;
            }

            private static bool? GetOptionalBoolean(OpenXmlLeafElement? element) {
                if (element == null) {
                    return null;
                }

                if (element is BooleanPropertyType booleanProperty && booleanProperty.Val != null) {
                    return booleanProperty.Val.Value;
                }

                return true;
            }
        }
    }
}
