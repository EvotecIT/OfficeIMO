using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Spreadsheet;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // The deterministic font estimates need the same allowance in both directions:
        // a row must wrap at the width that column autofit reserves for its text.
        private const double AutoFitTextWidthSafetyFactor = 1.22D;

        /// <summary>Resolves display inheritance without changing the cell or treating an explicit XF zero as absent.</summary>
        private static uint ResolveDisplayStyleIndex(uint? explicitStyle, Row? row, int column) {
            if (explicitStyle.HasValue) return explicitStyle.Value;
            if (row?.Parent is not SheetData sheetData) return 0U;
            return ResolveInheritedCellStyle(sheetData, row, column);
        }

        private static uint GetCellDisplayStyleIndex(Cell cell) {
            if (cell.StyleIndex?.Value is uint style) return style;
            string? reference = cell.CellReference?.Value;
            return string.IsNullOrEmpty(reference) ? 0U
                : ResolveDisplayStyleIndex(null, cell.Parent as Row, GetColumnIndex(reference!));
        }

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

        private AutoFitTextContext CreateAutoFitTextContext(bool includeSharedStrings = true) {
            var stylesheet = WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet;
            return new AutoFitTextContext(
                includeSharedStrings ? _excelDocument.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList() : null,
                stylesheet?.CellFormats?.Elements<CellFormat>().ToList(),
                BuildCustomNumberFormatCache(stylesheet));
        }

        private static Dictionary<uint, string?>? BuildCustomNumberFormatCache(Stylesheet? stylesheet) {
            var numberingFormats = stylesheet?.NumberingFormats;
            if (numberingFormats == null) {
                return null;
            }

            var cache = new Dictionary<uint, string?>();
            foreach (var format in numberingFormats.Elements<NumberingFormat>()) {
                uint id = format.NumberFormatId?.Value ?? 0U;
                cache[id] = format.FormatCode?.Value;
            }

            return cache;
        }

        private string GetCellAutoFitText(Cell cell, AutoFitTextContext? context = null) =>
            GetCellAutoFitText(cell, context ?? CreateAutoFitTextContext(cell.DataType?.Value == XCellValues.SharedString), _excelDocument.DateSystem);

        private static string GetCellAutoFitText(Cell cell, AutoFitTextContext context, ExcelDateSystem dateSystem) {
            var dataType = cell.DataType?.Value;

            if (dataType == XCellValues.SharedString) {
                return FormatAutoFitTextValue(cell, GetCachedSharedStringText(cell, context), context);
            }

            if (dataType == XCellValues.InlineString) {
                return FormatAutoFitTextValue(cell, GetInlineStringText(cell.InlineString), context);
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

            if (dataType == XCellValues.Date
                && DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date)) {
                uint numberFormatId = GetCellNumberFormatId(cell, context);
                string? formatCode = GetNumberFormatCode(numberFormatId, context);
                if (!string.IsNullOrWhiteSpace(formatCode) && !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                    return ExcelNumberFormatDisplay.FormatDateTimeText(date, numberFormatId, formatCode!,
                        ExcelDateSystemConverter.ToSerialForDisplay(date, dateSystem)) ?? raw;
                }
            }

            if (dataType == null || dataType == XCellValues.Number) {
                return FormatAutoFitNumericText(cell, raw, context, dateSystem);
            }

            return dataType == XCellValues.String ? FormatAutoFitTextValue(cell, Utilities.ExcelSanitizer.DecodeXString(raw), context) : raw;
        }

        private static string FormatAutoFitTextValue(Cell cell, string text, AutoFitTextContext context) =>
            FormatAutoFitTextValue(text, GetCellDisplayStyleIndex(cell), context);

        private static string FormatAutoFitTextValue(string text, uint styleIndex, AutoFitTextContext context) =>
            SpreadsheetNumberFormatDisplay.FormatTextValue(text, GetNumberFormatCode(GetCellNumberFormatId(styleIndex, context), context));

        private static string FormatAutoFitNumericText(Cell cell, string raw, AutoFitTextContext context, ExcelDateSystem dateSystem) {
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                return raw;
            }

            uint numberFormatId = GetCellNumberFormatId(cell, context);
            if (numberFormatId == 0U) {
                return raw;
            }

            string? formatCode = GetNumberFormatCode(numberFormatId, context);

            return ExcelNumberFormatDisplay.FormatNumericText(value, numberFormatId, formatCode, raw, dateSystem);
        }

        private static uint GetCellNumberFormatId(Cell cell, AutoFitTextContext context) {
            return GetCellNumberFormatId(GetCellDisplayStyleIndex(cell), context);
        }

        private static uint GetCellNumberFormatId(uint styleIndex, AutoFitTextContext context) {
            if (context.NumberFormatIdsByStyle.TryGetValue(styleIndex, out uint cachedFormatId)) {
                return cachedFormatId;
            }

            var cellFormats = context.CellFormats;
            CellFormat? cellFormat = cellFormats != null && styleIndex < cellFormats.Count ? cellFormats[(int)styleIndex] : null;

            uint numberFormatId = cellFormat?.NumberFormatId?.Value ?? 0U;
            context.NumberFormatIdsByStyle[styleIndex] = numberFormatId;

            return numberFormatId;
        }

        private static string? GetNumberFormatCode(uint numberFormatId, AutoFitTextContext context) {
            if (context.NumberFormatCodes.TryGetValue(numberFormatId, out string? cachedCode)) {
                return cachedCode;
            }

            string? code = GetNumberFormatCodeCore(numberFormatId, context);
            context.NumberFormatCodes[numberFormatId] = code;

            return code;
        }

        private static string? GetNumberFormatCodeCore(uint numberFormatId, AutoFitTextContext context) {
            string? builtIn = ExcelBuiltInNumberFormats.GetCode(numberFormatId);
            if (builtIn != null) {
                return builtIn;
            }

            if (context.CustomNumberFormats != null) {
                return context.CustomNumberFormats.TryGetValue(numberFormatId, out string? formatCode) ? formatCode : null;
            }

            return null;
        }

        private static bool TryGetSharedStringIndex(Cell cell, out int id) {
            var raw = cell.CellValue?.InnerText;
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out id);
        }

        private static string GetCachedSharedStringText(Cell cell, AutoFitTextContext context) {
            if (!TryGetSharedStringIndex(cell, out int id)) {
                return string.Empty;
            }

            if (context.SharedStringTexts.TryGetValue(id, out string? cachedText)) {
                return cachedText;
            }

            var sharedStrings = context.SharedStrings;
            string text = sharedStrings != null && id >= 0 && id < sharedStrings.Count
                ? GetSharedStringText(sharedStrings[id])
                : string.Empty;

            context.SharedStringTexts[id] = text;

            return text;
        }

        private static IReadOnlyList<AutoFitTextRun>? GetCellAutoFitRichTextRuns(Cell cell, AutoFitTextContext context) {
            IReadOnlyList<AutoFitTextRun>? runs;
            if (cell.DataType?.Value == XCellValues.SharedString) {
                if (!TryGetSharedStringIndex(cell, out int id)) {
                    return null;
                }

                if (!context.SharedStringRichTextRuns.TryGetValue(id, out runs)) {
                    var sharedStrings = context.SharedStrings;
                    runs = sharedStrings != null && id >= 0 && id < sharedStrings.Count
                        ? GetSharedStringRichTextRuns(sharedStrings[id]) : null;
                    context.SharedStringRichTextRuns[id] = runs;
                }
            } else if (cell.DataType?.Value == XCellValues.InlineString) {
                runs = GetInlineStringRichTextRuns(cell.InlineString);
            } else {
                return null;
            }
            if (runs == null || runs.Count == 0) return runs;
            string? formatCode = GetNumberFormatCode(GetCellNumberFormatId(cell, context), context);
            if (string.IsNullOrEmpty(formatCode) || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)
                || string.Equals(formatCode, "@", StringComparison.Ordinal)) return runs;
            string raw = string.Concat(runs.Select(run => run.Text));
            var displayedRuns = new List<AutoFitTextRun>();
            SpreadsheetNumberFormatDisplay.ComposeTextValue(raw, formatCode, (text, isValue) => {
                if (isValue) displayedRuns.AddRange(runs);
                else if (text.Length > 0) displayedRuns.Add(AutoFitTextRun.Create(text, null));
            });
            return displayedRuns;
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
                string text = Utilities.ExcelSanitizer.DecodeXString(run.Text?.Text ?? string.Empty);
                if (text.Length == 0) {
                    continue;
                }

                result.Add(AutoFitTextRun.Create(text, run.RunProperties));
            }

            return result;
        }

        private static string GetSharedStringText(SharedStringItem item) {
            if (item.Text != null) {
                return Utilities.ExcelSanitizer.DecodeXString(item.Text.Text ?? string.Empty);
            }

            var sb = new StringBuilder();
            foreach (var text in item.Descendants<Text>()) {
                sb.Append(Utilities.ExcelSanitizer.DecodeXString(text.Text));
            }

            return sb.ToString();
        }

        private static string GetInlineStringText(InlineString? inlineString) {
            if (inlineString == null) {
                return string.Empty;
            }

            if (inlineString.Text != null) {
                return Utilities.ExcelSanitizer.DecodeXString(inlineString.Text.Text ?? string.Empty);
            }

            var sb = new StringBuilder();
            foreach (var run in inlineString.Elements<Run>()) {
                if (run.Text != null) {
                    sb.Append(Utilities.ExcelSanitizer.DecodeXString(run.Text.Text));
                }
            }

            return sb.ToString();
        }

        internal readonly struct AutoFitTextRun {
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
                Size.HasValue ? InvariantNumberText.Get(Size.Value) : string.Empty,
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
