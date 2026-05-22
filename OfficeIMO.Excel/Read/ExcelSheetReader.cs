using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Xml;

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
        private static readonly XmlReaderSettings WorksheetXmlReaderSettings = CreateWorksheetXmlReaderSettings();

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
            return CanUseEnumerateCellsXmlReader()
                ? EnumerateCellsXmlFast(CancellationToken.None)
                : EnumerateCellsDom(CancellationToken.None);
        }

        private IEnumerable<CellValueInfo> EnumerateCellsDom(CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = checked((int)row.RowIndex!.Value);
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    var value = ConvertCell(cell);
                    if (value is not null || CellHasExplicitBlank(cell))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }

        private IEnumerable<CellValueInfo> EnumerateCellsXmlFast(CancellationToken ct) {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            bool hasCustomConverter = _opt.CellValueConverter != null;
            int nextRowIndex = 1;

            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                    continue;
                }

                int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                if (rowIndex <= 0) {
                    rowIndex = nextRowIndex;
                }

                nextRowIndex = rowIndex + 1;
                if (reader.IsEmptyElement) {
                    continue;
                }

                int depth = reader.Depth;
                int nextColumnIndex = 1;
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "row") {
                        break;
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                        continue;
                    }

                    int columnIndex = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                    if (columnIndex <= 0) {
                        SkipXmlElement(reader, "c");
                        continue;
                    }

                    if (hasCustomConverter) {
                        if (TryReadXmlCellValueForCellEnumeration(reader, rowIndex, columnIndex, out object? customValue, out bool explicitBlank)) {
                            if (customValue != null || explicitBlank) {
                                yield return new CellValueInfo(rowIndex, columnIndex, customValue);
                            }
                        }

                        continue;
                    }

                    if (reader.IsEmptyElement) {
                        continue;
                    }

                    object? cellValue = ReadXmlCellValue(reader);
                    if (cellValue != null) {
                        yield return new CellValueInfo(rowIndex, columnIndex, cellValue);
                    }
                }
            }
        }

        private bool TryReadXmlCellValueForCellEnumeration(XmlReader cellReader, int rowIndex, int columnIndex, out object? value, out bool explicitBlank) {
            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            bool readStyleIndex = true;

            CellRaw raw = ReadXmlCellRaw(cellReader, rowIndex, columnIndex, cellKind, readStyleIndex);
            explicitBlank = raw.RawText != null && raw.RawText.Length == 0 && raw.InlineText == null && raw.FormulaText == null;
            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null) {
                value = null;
                return false;
            }

            value = ConvertRaw(raw).TypedValue;
            return true;
        }

        private bool CanUseEnumerateCellsXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
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
                RewindWorksheetStream(stream);
                return !stream.CanSeek || stream.Length > 0;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

        private static void RewindWorksheetStream(Stream stream) {
            if (stream.CanSeek) {
                stream.Position = 0;
            }
        }

        private static XmlReader OpenWorksheetXmlReader(Stream stream) {
            return XmlReader.Create(stream, WorksheetXmlReaderSettings);
        }

        private static XmlReaderSettings CreateWorksheetXmlReaderSettings() {
            return new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            };
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
                && _styles.HasDateStyles
                && typeHint != CellValues.SharedString
                && typeHint != CellValues.Boolean
                && typeHint != CellValues.String
                && typeHint != CellValues.InlineString
                && typeHint != CellValues.Date;
        }

        private static XmlCellKind ParseXmlCellKind(string? type) {
            if (string.IsNullOrEmpty(type)) {
                return XmlCellKind.Default;
            }

            string text = type!;
            switch (text.Length) {
                case 1:
                    return text[0] switch {
                        'b' => XmlCellKind.Boolean,
                        'd' => XmlCellKind.Date,
                        'n' => XmlCellKind.Number,
                        's' => XmlCellKind.SharedString,
                        _ => XmlCellKind.Unknown
                    };
                case 3:
                    return text == "str" ? XmlCellKind.String : XmlCellKind.Unknown;
                case 9:
                    return text == "inlineStr" ? XmlCellKind.InlineString : XmlCellKind.Unknown;
                default:
                    return XmlCellKind.Unknown;
            }
        }

        private static bool CellKindCanUseDateStyle(XmlCellKind kind) {
            return kind == XmlCellKind.Default
                || kind == XmlCellKind.Number
                || kind == XmlCellKind.Unknown;
        }

        private static CellValues? ToCellValueType(XmlCellKind kind) {
            return kind switch {
                XmlCellKind.Boolean => CellValues.Boolean,
                XmlCellKind.Date => CellValues.Date,
                XmlCellKind.InlineString => CellValues.InlineString,
                XmlCellKind.Number => CellValues.Number,
                XmlCellKind.SharedString => CellValues.SharedString,
                XmlCellKind.String => CellValues.String,
                _ => null
            };
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
                if (TryParseInvariantDoubleFast(rawText, out var oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    value = DateTime.FromOADate(oa);
                } else {
                    value = rawText;
                }

                return true;
            }

            if (_opt.NumericAsDecimal) {
                if (TryParseRawDecimal(rawText, out var dec)) {
                    value = dec;
                    return true;
                }

                if (TryParseRawDouble(rawText, out var dbl)) {
                    value = dbl;
                    return true;
                }

                value = rawText;
                return true;
            }

            if (TryParseRawDouble(rawText, out var num)) {
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

        private static bool TryParseInvariantDoubleFast(string? rawText, out double value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            string text = rawText!;
            int length = text.Length;
            int index = 0;
            bool negative = false;
            if (text[0] == '-') {
                negative = true;
                index = 1;
                if (index == length) {
                    return false;
                }
            } else if (text[0] == '+') {
                index = 1;
                if (index == length) {
                    return false;
                }
            }

            long whole = 0;
            bool hasDigit = false;
            for (; index < length; index++) {
                char ch = text[index];
                int digit = ch - '0';
                if ((uint)digit > 9U) {
                    break;
                }

                if (whole > (long.MaxValue - digit) / 10) {
                    return false;
                }

                whole = (whole * 10) + digit;
                hasDigit = true;
            }

            double parsed = whole;
            if (index < length && text[index] == '.') {
                index++;
                double scale = 0.1D;
                for (; index < length; index++) {
                    char ch = text[index];
                    int digit = ch - '0';
                    if ((uint)digit > 9U) {
                        break;
                    }

                    parsed += digit * scale;
                    scale *= 0.1D;
                    hasDigit = true;
                }
            }

            if (!hasDigit || index != length) {
                return false;
            }

            value = negative ? -parsed : parsed;
            return true;
        }

        private bool TryParseRawDouble(string rawText, out double value) {
            if (_opt.Culture != CultureInfo.InvariantCulture
                && double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out value)) {
                return true;
            }

            return TryParseInvariantDoubleFast(rawText, out value)
                || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private bool TryParseRawInt32(string rawText, out int value) {
            if (_opt.Culture == CultureInfo.InvariantCulture && TryParseInvariantInt32Fast(rawText, out value)) {
                return true;
            }

            return int.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out value);
        }

        private static bool TryParseInvariantInt32Fast(string rawText, out int value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            uint limit = negative ? 2147483648U : int.MaxValue;
            uint parsed = 0U;
            for (; index < rawText.Length; index++) {
                int digit = rawText[index] - '0';
                if ((uint)digit > 9U) {
                    return false;
                }

                if (parsed > (limit - (uint)digit) / 10U) {
                    return false;
                }

                parsed = (parsed * 10U) + (uint)digit;
            }

            if (negative) {
                value = parsed == 2147483648U ? int.MinValue : -(int)parsed;
            } else {
                value = (int)parsed;
            }

            return true;
        }

        private bool TryParseRawInt64(string rawText, out long value) {
            if (_opt.Culture == CultureInfo.InvariantCulture && TryParseInvariantInt64Fast(rawText, out value)) {
                return true;
            }

            return long.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out value);
        }

        private static bool TryParseInvariantInt64Fast(string rawText, out long value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            ulong limit = negative ? 9223372036854775808UL : long.MaxValue;
            ulong parsed = 0UL;
            for (; index < rawText.Length; index++) {
                int digit = rawText[index] - '0';
                if ((uint)digit > 9U) {
                    return false;
                }

                if (parsed > (limit - (uint)digit) / 10UL) {
                    return false;
                }

                parsed = (parsed * 10UL) + (uint)digit;
            }

            if (negative) {
                value = parsed == 9223372036854775808UL ? long.MinValue : -(long)parsed;
            } else {
                value = (long)parsed;
            }

            return true;
        }

        private bool TryParseRawDecimal(string rawText, out decimal value) {
            return TryParseRawDecimal(rawText, _opt.Culture, out value);
        }

        private static bool TryParseRawDecimal(string rawText, CultureInfo culture, out decimal value) {
            if (culture == CultureInfo.InvariantCulture) {
                return TryParseInvariantDecimalFast(rawText, out value)
                    || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
            }

            if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, culture, out value)) {
                return true;
            }

            return TryParseInvariantDecimalFast(rawText, out value)
                || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryParseInvariantDecimalFast(string rawText, out decimal value) {
            value = 0m;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            ulong parsed = 0UL;
            int scale = 0;
            bool hasDigit = false;
            bool hasDecimalPoint = false;
            for (; index < rawText.Length; index++) {
                char ch = rawText[index];
                int digit = ch - '0';
                if ((uint)digit <= 9U) {
                    if (parsed > (ulong.MaxValue - (uint)digit) / 10UL) {
                        return false;
                    }

                    parsed = (parsed * 10UL) + (uint)digit;
                    hasDigit = true;
                    if (hasDecimalPoint && ++scale > 28) {
                        return false;
                    }

                    continue;
                }

                if (ch == '.' && !hasDecimalPoint) {
                    hasDecimalPoint = true;
                    continue;
                }

                return false;
            }

            if (!hasDigit) {
                return false;
            }

            value = new decimal(
                (int)(parsed & 0xFFFFFFFF),
                (int)((parsed >> 32) & 0xFFFFFFFF),
                0,
                negative,
                (byte)scale);
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
                    if (TryParseInvariantDoubleFast(rawText, out var oa)
                        || double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        return System.DateTime.FromOADate(oa);
                }
                if (_opt.NumericAsDecimal) {
                    if (TryParseRawDecimal(rawText, out var dec))
                        return dec;
                    if (TryParseRawDouble(rawText, out var dbl))
                        return dbl;
                    return rawText;
                } else {
                    if (TryParseRawDouble(rawText, out var num))
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
                    if (TryParseInvariantDoubleFast(rawText, out var oa)
                        || double.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        return System.DateTime.FromOADate(oa);
                    return rawText;
                }

                if (_opt.NumericAsDecimal) {
                    if (TryParseRawDecimal(rawText, out var dec2))
                        return dec2;
                    if (TryParseRawDouble(rawText, out var dbl2))
                        return dbl2;
                } else {
                    if (TryParseRawDouble(rawText, out var num))
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

        private enum XmlCellKind {
            Default,
            Boolean,
            Date,
            InlineString,
            Number,
            SharedString,
            String,
            Unknown
        }
    }
}
