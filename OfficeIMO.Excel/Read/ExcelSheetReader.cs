using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reader for a single worksheet. Offers enumeration and conversion helpers.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private readonly string _sheetName;
        private readonly WorksheetPart _wsPart;
        private readonly SharedStringCache _sst;
        private readonly StylesCache _styles;
        private readonly ExcelReadOptions _opt;
        private readonly bool _canStreamWorksheetPart;
        private bool? _hasWorksheetPartStreamContent;

        internal ExcelSheetReader(string sheetName, WorksheetPart wsPart, SharedStringCache sst, StylesCache styles, ExcelReadOptions opt, bool canStreamWorksheetPart) {
            _sheetName = sheetName;
            _wsPart = wsPart;
            _sst = sst;
            _styles = styles;
            _opt = opt;
            _canStreamWorksheetPart = canStreamWorksheetPart;
        }

        /// <summary>
        /// Worksheet name.
        /// </summary>
        public string Name => _sheetName;

        /// <summary>
        /// Enumerates non-empty cells as (Row, Column, Value). Values are typed when possible.
        /// </summary>
        public IEnumerable<CellValueInfo> EnumerateCells() {
            foreach (var row in EnumerateWorksheetRows()) {
                var rIndex = checked((int)row.RowIndex!.Value);
                foreach (var cell in row.Elements<Cell>()) {
                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    var value = ConvertCell(cell);
                    if (value is not null || CellHasExplicitBlank(cell))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }

        // ---------- Internals ----------

        private static bool CellHasExplicitBlank(Cell cell) {
            return cell.CellValue is not null && string.IsNullOrEmpty(cell.CellValue.InnerText);
        }

        private static string? ExtractRawText(Cell cell) {
            return cell.CellValue?.InnerText;
        }

        private IEnumerable<Row> EnumerateWorksheetRows(CancellationToken ct = default) {
            if (CanStreamWorksheetPart()) {
                foreach (var row in EnumerateWorksheetRowsFromPart(ct)) {
                    yield return row;
                }

                yield break;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                yield break;
            }

            bool canCancel = ct.CanBeCanceled;
            foreach (var row in sheetData.Elements<Row>()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                yield return row;
            }
        }

        private IEnumerable<Row> EnumerateWorksheetRowsFromPart(CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            using var reader = OpenXmlReader.Create(_wsPart);
            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (reader.IsStartElement && reader.ElementType == typeof(Row)) {
                    if (reader.LoadCurrentElement() is Row row) {
                        yield return row;
                    }
                }
            }
        }

        private bool CanStreamWorksheetPart() {
            if (!_canStreamWorksheetPart) {
                return false;
            }

            if (_hasWorksheetPartStreamContent is bool hasContent) {
                return hasContent;
            }

            bool result = HasWorksheetPartStreamContent();
            _hasWorksheetPartStreamContent = result;
            return result;
        }

        private bool HasWorksheetPartStreamContent() {
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                return !stream.CanSeek || stream.Length > 0;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

        private static string? ExtractFormulaText(Cell cell) {
            return cell.CellFormula?.Text;
        }

        private static string? ExtractInlineString(Cell cell, CellValues? typeHint) {
            if (typeHint != CellValues.InlineString && cell.InlineString is null) return null;

            var inline = cell.InlineString;
            if (inline?.Text?.Text != null) return inline.Text.Text;
            if (inline?.HasChildren == true) {
                return SharedStringCache.GetRunText(inline);
            }
            return null;
        }

        private object? ConvertCell(Cell cell) {
            TryConvertCell(cell, out var value);
            return value;
        }

        private bool TryConvertCell(Cell cell, out object? value) {
            value = null;
            CellValues? typeHint = cell.DataType?.Value;
            bool hasFormula = cell.CellFormula is not null;
            string? formulaText = null;
            if (hasFormula) {
                formulaText = ExtractFormulaText(cell);
                if (!_opt.UseCachedFormulaResult && formulaText != null) {
                    value = formulaText;
                    return true;
                }
            }

            string? rawText = ExtractRawText(cell);
            string? inlineText = ExtractInlineString(cell, typeHint);
            if (hasFormula && _opt.UseCachedFormulaResult && rawText == null && formulaText != null) {
                value = formulaText;
                return true;
            }

            if (rawText == null && inlineText == null && formulaText == null) {
                if (!CellHasExplicitBlank(cell) && !_opt.FillBlanksInRanges) {
                    return false;
                }
            }

            if (hasFormula && !_opt.UseCachedFormulaResult) {
                value = formulaText ?? rawText ?? inlineText;
                return true;
            }

            uint? styleIndex = null;
            if (NeedsStyleForConversion(typeHint, rawText)) {
                styleIndex = cell.StyleIndex?.Value;
            }

            value = TryConvertWithoutCustomHook(typeHint, styleIndex, rawText, inlineText, out object? converted)
                ? converted
                : ConvertByHints(typeHint, styleIndex ?? cell.StyleIndex?.Value, rawText, inlineText);
            return true;
        }

        private bool NeedsStyleForConversion(CellValues? typeHint, string? rawText) {
            return rawText != null
                && _opt.TreatDatesUsingNumberFormat
                && typeHint != CellValues.SharedString
                && typeHint != CellValues.Boolean
                && typeHint != CellValues.String
                && typeHint != CellValues.InlineString
                && typeHint != CellValues.Date;
        }

        private CellRaw SnapshotCell(Cell cell, int row = 0, int col = 0) {
            var hasFormula = cell.CellFormula is not null;
            var formulaText = hasFormula ? ExtractFormulaText(cell) : null;
            var preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            var typeHint = cell.DataType?.Value;

            return new CellRaw {
                Row = row,
                Col = col,
                TypeHint = typeHint,
                StyleIndex = cell.StyleIndex?.Value,
                HasFormula = hasFormula,
                FormulaText = formulaText,
                RawText = preferFormulaText ? null : ExtractRawText(cell),
                InlineText = preferFormulaText ? null : ExtractInlineString(cell, typeHint)
            };
        }

        private CellRaw ConvertRaw(CellRaw raw) {
            if (raw.HasFormula) {
                if (_opt.UseCachedFormulaResult && raw.RawText != null) {
                    raw.TypedValue = TryConvertWithoutCustomHook(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out var cachedValue)
                        ? cachedValue
                        : ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText);
                } else {
                    raw.TypedValue = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                }
                return raw;
            }

            raw.TypedValue = TryConvertWithoutCustomHook(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out var value)
                ? value
                : ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText);
            return raw;
        }

        private bool TryConvertWithoutCustomHook(CellValues? type, uint? styleIndex, string? rawText, string? inlineText, out object? value) {
            value = null;
            if (_opt.CellValueConverter != null) {
                return false;
            }

            if (!string.IsNullOrEmpty(inlineText)) {
                value = inlineText;
                return true;
            }

            if (type == CellValues.SharedString) {
                if (TryParseSharedStringIndex(rawText, out var sstIndex)) {
                    value = _sst.Get(sstIndex);
                    return true;
                }

                value = rawText;
                return rawText != null;
            }

            if (type == CellValues.Boolean && rawText != null) {
                value = rawText == "1";
                return true;
            }

            if (type == CellValues.String || type == CellValues.InlineString) {
                value = rawText ?? inlineText;
                return value != null;
            }

            if (type == CellValues.Date && rawText != null) {
                if (DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                    value = dt;
                } else {
                    value = rawText;
                }

                return true;
            }

            if (rawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && _styles.IsDateLike(styleIndex.Value)) {
                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa)) {
                    value = DateTime.FromOADate(oa);
                } else {
                    value = rawText;
                }

                return true;
            }

            if (_opt.NumericAsDecimal) {
                if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out var dec)) {
                    value = dec;
                    return true;
                }

                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out var dbl)) {
                    value = dbl;
                    return true;
                }

                value = rawText;
                return true;
            }

            if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out var num)) {
                value = num;
            } else {
                value = rawText;
            }

            return true;
        }

        private static bool TryParseSharedStringIndex(string? rawText, out int index) {
            index = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            string text = rawText!;
            int parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
            }

            index = parsed;
            return true;
        }

        private object? ConvertByHints(CellValues? type, uint? styleIndex, string? rawText, string? inlineText) {
            // Custom converter hook (cell-level). If provided and handled, honor it.
            var hook = _opt.CellValueConverter;
            if (hook != null) {
                var ctx = new ExcelCellContext(type, styleIndex, rawText, inlineText, _opt.Culture);
                var res = hook(ctx);
                if (res.Handled) return res.Value;
            }
            if (!string.IsNullOrEmpty(inlineText)) return inlineText;

            if (type == CellValues.SharedString && TryParseSharedStringIndex(rawText, out var sstIndex))
                return _sst.Get(sstIndex);

            if (type == CellValues.Boolean && rawText != null)
                return rawText == "1";

            if (type == CellValues.Number && rawText != null) {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && _styles.IsDateLike(styleIndex.Value)) {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa))
                        return System.DateTime.FromOADate(oa);
                }
                if (_opt.NumericAsDecimal) {
                    if (decimal.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var dec))
                        return dec;
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var dbl))
                        return dbl;
                    return rawText;
                } else {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var num))
                        return num;
                }
                return rawText;
            }

            if (type == CellValues.Date && rawText != null) {
                if (System.DateTime.TryParse(rawText, _opt.Culture, System.Globalization.DateTimeStyles.AssumeLocal, out var dt))
                    return dt;
                return rawText;
            }

            if (type == CellValues.String || type == CellValues.InlineString || type == CellValues.SharedString) {
                if (type == CellValues.SharedString && TryParseSharedStringIndex(rawText, out var idx))
                    return _sst.Get(idx);
                return rawText ?? inlineText;
            }

            if (rawText != null) {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && _styles.IsDateLike(styleIndex.Value)) {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var oa))
                        return System.DateTime.FromOADate(oa);
                    return rawText;
                }

                if (_opt.NumericAsDecimal) {
                    if (decimal.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var dec2))
                        return dec2;
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var dbl2))
                        return dbl2;
                } else {
                    if (double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, _opt.Culture, out var num))
                        return num;
                }
                return rawText;
            }

            return null;
        }

        private struct CellRaw {
            public int Row;
            public int Col;
            public CellValues? TypeHint;
            public uint? StyleIndex;
            public bool HasFormula;
            public string? FormulaText;
            public string? RawText;
            public string? InlineText;
            public object? TypedValue;
        }
    }
}
