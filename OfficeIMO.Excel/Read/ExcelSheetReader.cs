using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Data;
using System.Globalization;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reader for a single worksheet. Offers enumeration and conversion helpers.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        private const string SpreadsheetNamespace =
            SpreadsheetXmlContent.SpreadsheetNamespace;
        private const string StrictSpreadsheetNamespace =
            SpreadsheetXmlContent.StrictSpreadsheetNamespace;
        private readonly string _sheetName;
        private const int SpreadsheetXml = OfficeXmlNamespaces.SpreadsheetML;
        private readonly ExcelPartRef _worksheetPart;
        private readonly bool _hasLiveWorksheetPart;
        private readonly ExcelReaderPackage? _package;
        private readonly bool _readOnlySnapshot;
        private readonly string _worksheetPartName;
        private readonly bool _allowGeneralProjection;
        private readonly SharedStringCache _sst;
        private readonly StylesCacheProvider _styles;
        private readonly ExcelReaderOptions _opt;
        private readonly ExcelDateSystem _dateSystem;
        private readonly bool _canStreamWorksheetPart;
        private readonly OpenXmlPackagePartBufferReader? _partBufferReader;
        private readonly Lazy<RichValueErrorLookup> _richValueErrors;
        private StylesCache? _stylesCache;
        private OfficeXmlNode _packageWorksheetRoot;
        private bool? _hasWorksheetPartStreamContent;
        private string? _usedRangeA1;
        private char[]? _xmlValueTextBuffer;
        private static readonly XmlReaderSettings WorksheetXmlReaderSettings = CreateWorksheetXmlReaderSettings();
        private static readonly XmlReaderSettings CellEnumerationXmlReaderSettings = CreateCellEnumerationXmlReaderSettings();
        private static readonly object BoxedTrue = true;
        private static readonly object BoxedFalse = false;

        internal ExcelSheetReader(
            string sheetName,
            ExcelPartRef worksheetPart,
            SharedStringCache sst,
            StylesCacheProvider styles,
            ExcelReaderOptions opt,
            ExcelDateSystem dateSystem,
            bool canStreamWorksheetPart,
            OpenXmlPackagePartBufferReader? partBufferReader,
            Lazy<RichValueErrorLookup> richValueErrors) {
            _sheetName = sheetName;
            _worksheetPart = worksheetPart;
            _hasLiveWorksheetPart = true;
            _worksheetPartName = worksheetPart.Name;
            _allowGeneralProjection = true;
            _sst = sst;
            _styles = styles;
            _opt = opt;
            _dateSystem = dateSystem;
            _canStreamWorksheetPart = canStreamWorksheetPart;
            _partBufferReader = partBufferReader;
            _richValueErrors = richValueErrors;
        }

        internal ExcelSheetReader(
            string sheetName,
            string worksheetPartName,
            SharedStringCache sst,
            StylesCacheProvider styles,
            ExcelReaderOptions opt,
            ExcelDateSystem dateSystem,
            OpenXmlPackagePartBufferReader partBufferReader,
            Lazy<RichValueErrorLookup>? richValueErrors = null,
            ExcelReaderPackage? package = null,
            bool canStreamWorksheetPart = true,
            bool readOnlySnapshot = false) {
            _sheetName = sheetName;
            _worksheetPartName = worksheetPartName;
            _allowGeneralProjection = package != null;
            _package = package;
            _readOnlySnapshot = readOnlySnapshot;
            _sst = sst;
            _styles = styles;
            _opt = opt;
            _dateSystem = dateSystem;
            _canStreamWorksheetPart = canStreamWorksheetPart;
            _partBufferReader = partBufferReader;
            _richValueErrors = richValueErrors ?? new Lazy<RichValueErrorLookup>(() => RichValueErrorLookup.Empty);
            _hasWorksheetPartStreamContent = true;
        }

        private OfficeXmlNode WorksheetRoot {
            get {
                OfficeXmlNode root;
                if (_hasLiveWorksheetPart) {
                    root = _worksheetPart.Root;
                } else {
                    if (_packageWorksheetRoot.Exists) return _packageWorksheetRoot;
                    using Stream stream = OpenReaderWorksheetStream(_opt.CancellationToken);
                    root = ExcelReaderPackage.ReadPartRoot(stream, _opt.CancellationToken, validateMarkup: true);
                }
                if (!root.Is(SpreadsheetXml, "worksheet")) throw new InvalidOperationException("Worksheet is null.");
                if (!_hasLiveWorksheetPart) _packageWorksheetRoot = root;
                return root;
            }
        }

        private void RequireGeneralProjection() {
            if (!_allowGeneralProjection) {
                throw new XlsxTabularFastPathNotSupportedException(
                    $"Worksheet '{_sheetName}' requires the general worksheet projection.");
            }
        }

        // Native tabular readers retain the opened package snapshot. A worksheet
        // that is too large to index can still use the existing XML projection.
        private Stream OpenReaderWorksheetStream(CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            if (_partBufferReader != null
                && _partBufferReader.TryGetLength(_worksheetPartName, out long length)
                && length >= 0 && length <= int.MaxValue) {
                return _partBufferReader.OpenPart(_worksheetPartName, int.MaxValue, ct);
            }
            if (_hasLiveWorksheetPart) return _worksheetPart.OpenRead();
            throw new InvalidDataException($"Worksheet '{_sheetName}' has no readable worksheet part.");
        }

        private DateTime FromExcelSerialDate(double serial, bool calendarStyle) => calendarStyle ? ExcelDateSystemConverter.FromSerial(serial, _dateSystem) : ExcelDateSystemConverter.FromElapsedSerial(serial);
        private DateTime FromExcelSerialDate(double serial, uint? styleIndex) => FromExcelSerialDate(serial, styleIndex.HasValue && Styles.IsDateSystemShiftStyle(styleIndex.Value));
        private bool TryFromExcelSerialDate(double serial, bool calendarStyle, out DateTime value) =>
            ExcelDateSystemConverter.TryFromSerial(serial, _dateSystem, calendarStyle, out value);
        /// <summary>Typed date access: a serial without a DateTime is an invalid cast, as for any other non-date value.</summary>
        private DateTime DateOrInvalidCast(double serial, bool calendarStyle) =>
            TryFromExcelSerialDate(serial, calendarStyle, out DateTime value)
                ? value
                : throw new InvalidCastException($"The date-formatted value {serial.ToString("R", CultureInfo.InvariantCulture)} has no DateTime representation.");
        private object DateOrSerial(double serial, bool calendarStyle) => ExcelDateSystemConverter.FromSerialOrNumber(serial, _dateSystem, calendarStyle);
        private object DateOrSerial(double serial, uint? styleIndex) => DateOrSerial(serial, styleIndex.HasValue && Styles.IsDateSystemShiftStyle(styleIndex.Value));

        private StylesCache Styles => _stylesCache ??= _styles.GetValue(_opt.CancellationToken);

        // Background indexing helpers may only read styles that this reader already resolved.
        private bool StylesLoaded => Volatile.Read(ref _stylesCache) != null;

        private static object BoxBoolean(bool value) => value ? BoxedTrue : BoxedFalse;

        private string? GetSharedString(int index) => _sst.Get(index);

        private bool TryGetSharedStringUtf8(int index, out ArraySegment<byte> value) =>
            _sst.TryGetUtf8(index, out value);

        /// <summary>
        /// Worksheet name.
        /// </summary>
        public string Name => _sheetName;

        /// <summary>
        /// Enumerates non-empty cells as (ReaderRow, Column, Value). Values are typed when possible.
        /// </summary>
        public IEnumerable<ExcelCellValueInfo> EnumerateCells(CancellationToken ct = default) {
            return CanUseEnumerateCellsXmlReader()
                ? EnumerateCellsStreamed(ct)
                : EnumerateCellsDom(ct);
        }

        /// <summary>
        /// Chooses the route when enumeration starts and hands out that route's own enumerator, so each cell crosses
        /// one iterator instead of a forwarding one as well.
        /// </summary>
        private IEnumerable<ExcelCellValueInfo> EnumerateCellsStreamed(CancellationToken ct) =>
            new DeferredCells(this, ct);

        private IEnumerable<ExcelCellValueInfo> CreateStreamedCells(CancellationToken ct) =>
            TryCreateEnumerateCellsUtf8Source(ct, out ExcelUtf8RangeRowSource? source,
                    out int firstRow, out int lastRow, out int firstOrdinal, out int lastOrdinal, out int declaredFirstColumn)
                ? EnumerateCellsUtf8(source!, firstRow, lastRow, firstOrdinal, lastOrdinal, declaredFirstColumn, ct)
                : EnumerateCellsXmlFast(ct);

        private sealed class DeferredCells : IEnumerable<ExcelCellValueInfo> {
            private readonly ExcelSheetReader _owner;
            private readonly CancellationToken _cancellationToken;

            internal DeferredCells(ExcelSheetReader owner, CancellationToken cancellationToken) {
                _owner = owner;
                _cancellationToken = cancellationToken;
            }

            public IEnumerator<ExcelCellValueInfo> GetEnumerator() =>
                _owner.CreateStreamedCells(_cancellationToken).GetEnumerator();

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private IEnumerable<ExcelCellValueInfo> EnumerateCellsDom(CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = row.Index;
                int nextDomColumnIndex = 1;
                foreach (var cell in row.Cells()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = cell.GetColumnIndex(ref nextDomColumnIndex);
                    if (cIndex <= 0) continue;
                    var value = ConvertCell(cell);
                    if (value is not null || CellHasExplicitBlank(cell))
                        yield return new ExcelCellValueInfo(rIndex, cIndex, value);
                }
            }
        }

        private IEnumerable<ExcelCellValueInfo> EnumerateCellsXmlFast(CancellationToken ct) {
            using var stream = OpenReaderWorksheetStream(default);
            RewindWorksheetStream(stream);
            using var reader = OpenCellEnumerationXmlReader(stream);
            var worksheetRows = new WorksheetXmlRowSelector();
            bool canCancel = ct.CanBeCanceled;
            // A culture also changes how text parses as a number, which only the raw conversion honours.
            bool hasCustomConverter = _opt.CellValueConverter != null
                || !ReferenceEquals(_opt.Culture, System.Globalization.CultureInfo.InvariantCulture);
            int nextRowIndex = 1;

            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (!worksheetRows.IsRowElement(reader)) {
                    continue;
                }

                int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text);
                if (rowIndex <= 0) {
                    rowIndex = ResolveImplicitXmlRowIndex(reader, nextRowIndex, ct);
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

                    if (!SpreadsheetXmlContent.IsDirectChildElement(reader, depth, "c")) {
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
                                yield return new ExcelCellValueInfo(rowIndex, columnIndex, customValue);
                            }
                        }

                        continue;
                    }

                    string? cellType = ReadValidatedXmlCellType(reader);
                    if (reader.IsEmptyElement && cellType != "e") {
                        continue;
                    }

                    object? cellValue = ReadXmlCellValue(reader, cellType);
                    if (cellValue != null) {
                        yield return new ExcelCellValueInfo(rowIndex, columnIndex, cellValue);
                    }
                }
            }
        }

        /// <summary>Rejects a cell type the SDK model rejects, which the tolerant XML projection would read as a number.</summary>
        private static string? ReadValidatedXmlCellType(XmlReader cellReader) {
            string? cellType = ReadXmlCellTypeAttribute(cellReader);
            return IsKnownXmlCellType(cellType)
                ? cellType
                : throw new FormatException($"The cell type '{cellType}' is not valid.");
        }

        private static bool IsKnownXmlCellType(string? cellType) =>
            cellType is null or "n" or "s" or "str" or "b" or "e" or "d" or "inlineStr";

        private bool TryReadXmlCellValueForCellEnumeration(XmlReader cellReader, int rowIndex, int columnIndex, out object? value, out bool explicitBlank) {
            XmlCellKind cellKind = ParseXmlCellKind(ReadValidatedXmlCellType(cellReader));
            bool readStyleIndex = true;

            CellRaw raw = ReadXmlCellRaw(cellReader, rowIndex, columnIndex, cellKind, readStyleIndex);
            explicitBlank = raw.RawText != null && raw.RawText.Length == 0 && raw.InlineText == null && raw.FormulaText == null;
            value = ConvertRaw(raw).TypedValue;
            return true;
        }

        private bool CanUseEnumerateCellsXmlReader() {
            return CanStreamWorksheetPart();
        }

        // ---------- Internals ----------

        private static bool CellHasExplicitBlank(ReaderCell cell) =>
            cell.Value is { Length: 0 };

        private IEnumerable<ReaderRow> EnumerateWorksheetRows(CancellationToken ct = default) {
            if (CanStreamWorksheetPart()) {
                return EnumerateWorksheetRowsFromPart(ct);
            }

            return EnumerateRowsWithCoordinates(WorksheetRoot.Child(SpreadsheetXml, "sheetData"), ct);
        }

        private IEnumerable<ReaderRow> EnumerateWorksheetRowsFromPart(CancellationToken ct) {
            using Stream stream = OpenReaderWorksheetStream(ct);
            RewindWorksheetStream(stream);
            using XmlReader reader = OpenCellEnumerationXmlReader(stream);
            var rows = new WorksheetXmlRowSelector();
            var namespaces = new OfficeXmlNamespaces.Table();
            int previousIndex = 0;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (!rows.IsRowElement(reader)) continue;

                using XmlReader fragment = reader.ReadSubtree();
                fragment.MoveToContent();
                byte[] xml = Encoding.UTF8.GetBytes(fragment.ReadOuterXml());
                OfficeXmlTree tree = OfficeXmlTree.Parse(xml, xml.Length, namespaces);
                var row = new OfficeXmlNode(tree, tree.Root);
                yield return new ReaderRow(row, ResolveRowIndex(row, ref previousIndex));
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
                using var stream = OpenReaderWorksheetStream(default);
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

        private static bool TryPrepareWorksheetStream(Stream stream) {
            if (!stream.CanSeek) {
                return true;
            }

            if (stream.Length == 0) {
                return false;
            }

            stream.Position = 0;
            return true;
        }

        private static XmlReader OpenWorksheetXmlReader(Stream stream) {
            return XmlReader.Create(stream, OpenXmlReadNameTable.WithSchemaNames(WorksheetXmlReaderSettings));
        }

        // A value such as <v>   </v> is cell text, so cell enumeration keeps whitespace nodes.
        private static XmlReader OpenCellEnumerationXmlReader(Stream stream) {
            return XmlReader.Create(stream, OpenXmlReadNameTable.WithSchemaNames(CellEnumerationXmlReaderSettings));
        }

        private static XmlReaderSettings CreateCellEnumerationXmlReaderSettings() {
            XmlReaderSettings settings = CreateWorksheetXmlReaderSettings();
            settings.IgnoreWhitespace = false;
            return settings;
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

        private static string? ExtractInlineString(OfficeXmlNode inline) {
            if (inline.IsNone) return null;
            OfficeXmlNode text = inline.Child(SpreadsheetXml, "t");
            if (text.Exists) return Utilities.ExcelSanitizer.DecodeXString(text.InnerText);
            return inline.FirstChild.Exists ? SharedStringCache.GetRunText(inline) : null;
        }

        private object? ConvertCell(ReaderCell cell) {
            TryConvertCell(cell, out var value);
            return value;
        }

        /// <summary>Converts a cell already loaded by a caller inspecting the live workbook.</summary>
        internal object? ConvertLoadedCell(object cell) => ConvertCell(new ReaderCell(OfficeXmlNode.FromSdk(cell)));

        private bool TryConvertCell(ReaderCell cell, out object? value) {
            value = null;
            ExcelCellValueType? typeHint = cell.Type;
            cell.GetChildren(out string? formula, out string? cellValue, out OfficeXmlNode inlineString);
            bool hasFormula = formula is not null;
            string? formulaText = null;
            if (hasFormula) {
                formulaText = formula;
                if (!_opt.UseCachedFormulaResult && formulaText != null) {
                    value = formulaText;
                    return true;
                }
            }

            string? cellText = cellValue;
            string? rawText = typeHint == ExcelCellValueType.Error && cell.ValueMetaIndex != null
                ? _richValueErrors.Value.Resolve(cell.ValueMetaIndex.Value, cellText)
                : cellText;
            string? inlineText = ExtractInlineString(inlineString);
            if (hasFormula && _opt.UseCachedFormulaResult && rawText == null && formulaText != null) {
                value = formulaText;
                return true;
            }

            if (rawText == null && inlineText == null && formulaText == null) {
                if (!(cellValue is not null && cellText!.Length == 0) && !_opt.FillBlanksInRanges) {
                    return false;
                }
            }

            if (hasFormula && !_opt.UseCachedFormulaResult) {
                value = formulaText ?? rawText ?? inlineText;
                return true;
            }

            uint? styleIndex = cell.Style;
            uint? displayStyle = styleIndex;
            if (!displayStyle.HasValue && _opt.DisplayText != null) {
                displayStyle = ResolveInheritedDisplayStyle(cell.RowStyle, GetReferenceColumn(cell.Reference));
            }
            if (_opt.CellValueConverter == null && TryFormatDisplayText(typeHint, displayStyle, rawText, out string displayed)) {
                value = displayed;
                return true;
            }
            if (!NeedsStyleForConversion(typeHint, rawText, styleIndex)) {
                styleIndex = null;
            }

            value = TryConvertWithoutCustomHook(typeHint, styleIndex, rawText, inlineText, out object? converted)
                ? converted
                : ConvertByHints(typeHint, styleIndex ?? cell.Style, rawText, inlineText);
            return true;
        }

        /// <summary>Applies displayed-text mode to a numeric or ISO date cell whose style has a number format.</summary>
        private bool TryFormatDisplayText(ExcelCellValueType? type, uint? styleIndex, string? rawText, out string text) {
            text = string.Empty;
            ExcelCellDisplayText? display = _opt.DisplayText;
            if (display == null || rawText == null || !display.IsFormatted(styleIndex)) return false;
            bool numeric = type == null || type == ExcelCellValueType.Number;
            return display.TryFormatRaw(numeric, type == ExcelCellValueType.Date, rawText, styleIndex, out text);
        }

        private bool NeedsStyleForConversion(ExcelCellValueType? typeHint, string? rawText, uint? styleIndex) {
            return rawText != null
                && styleIndex is not null
                && _opt.TreatDatesUsingNumberFormat
                && typeHint != ExcelCellValueType.SharedString
                && typeHint != ExcelCellValueType.Boolean
                && typeHint != ExcelCellValueType.String
                && typeHint != ExcelCellValueType.InlineString
                && typeHint != ExcelCellValueType.Date
                && Styles.HasDateStyles;
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
                        'e' => XmlCellKind.Error,
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

        private static ExcelCellValueType? ToCellValueType(XmlCellKind kind) {
            return kind switch {
                XmlCellKind.Boolean => ExcelCellValueType.Boolean,
                XmlCellKind.Date => ExcelCellValueType.Date,
                XmlCellKind.Error => ExcelCellValueType.Error,
                XmlCellKind.InlineString => ExcelCellValueType.InlineString,
                XmlCellKind.Number => ExcelCellValueType.Number,
                XmlCellKind.SharedString => ExcelCellValueType.SharedString,
                XmlCellKind.String => ExcelCellValueType.String,
                _ => null
            };
        }

        private CellRaw SnapshotCell(ReaderCell cell, int row = 0, int col = 0) {
            cell.GetChildren(out string? formula, out string? cellValue, out OfficeXmlNode inlineString);
            var hasFormula = formula is not null;
            var formulaText = formula;
            var preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            var typeHint = cell.Type;

            return new CellRaw {
                Row = row,
                Col = col,
                TypeHint = typeHint,
                StyleIndex = cell.Style,
                HasFormula = hasFormula,
                FormulaText = formulaText,
                RawText = preferFormulaText ? null : typeHint == ExcelCellValueType.Error && cell.ValueMetaIndex != null
                    ? _richValueErrors.Value.Resolve(cell.ValueMetaIndex.Value, cellValue) : cellValue,
                InlineText = preferFormulaText ? null : ExtractInlineString(inlineString)
            };
        }

        private CellRaw ConvertRaw(CellRaw raw) {
            raw.CustomValueHandled = false;
            if (_opt.CellValueConverter == null && (!raw.HasFormula || _opt.UseCachedFormulaResult)
                && TryFormatDisplayText(raw.TypeHint, raw.StyleIndex, raw.RawText, out string displayed)) {
                raw.TypedValue = displayed;
                return raw;
            }
            if (raw.HasFormula) {
                if (_opt.UseCachedFormulaResult && raw.RawText != null) {
                    raw.TypedValue = TryConvertWithoutCustomHook(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out var cachedValue)
                        ? cachedValue
                        : ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out raw.CustomValueHandled);
                } else {
                    raw.TypedValue = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                }
                return raw;
            }

            raw.TypedValue = TryConvertWithoutCustomHook(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out var value)
                ? value
                : ConvertByHints(raw.TypeHint, raw.StyleIndex, raw.RawText, raw.InlineText, out raw.CustomValueHandled);
            return raw;
        }

        private bool TryConvertWithoutCustomHook(ExcelCellValueType? type, uint? styleIndex, string? rawText, string? inlineText, out object? value) {
            value = null;
            if (_opt.CellValueConverter != null) {
                return false;
            }

            if (!string.IsNullOrEmpty(inlineText)) {
                value = inlineText;
                return true;
            }

            if (type == ExcelCellValueType.SharedString) {
                if (TryParseSharedStringIndex(rawText, out var sstIndex)) {
                    value = GetSharedString(sstIndex);
                    return true;
                }

                value = rawText;
                return rawText != null;
            }

            if (type == ExcelCellValueType.Boolean && rawText != null) {
                value = BoxBoolean(ExcelBooleanCellText.IsTrue(rawText));
                return true;
            }

            if (type == ExcelCellValueType.String || type == ExcelCellValueType.InlineString) {
                value = rawText != null ? Utilities.ExcelSanitizer.DecodeXString(rawText) : inlineText;
                return value != null;
            }

            if (type == ExcelCellValueType.Date && rawText != null) {
                if (ExcelIsoDateText.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                    value = dt;
                } else {
                    value = rawText;
                }

                return true;
            }

            if (rawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && Styles.IsDateLike(styleIndex.Value)) {
                if (TryParseInvariantDouble(rawText, out var oa)
                    || OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    value = DateOrSerial(oa, styleIndex);
                } else {
                    value = rawText;
                }

                return true;
            }

            if (_opt.NumericAsDecimal) {
                if (TryParseExcelNumberAsDecimal(rawText, _opt.Culture, out var dec)) {
                    value = dec;
                    return true;
                }

                if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var dbl)) {
                    value = dbl;
                    return true;
                }

                value = rawText;
                return true;
            }

            if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var num)) {
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

            return TryParseSharedStringIndex(rawText.AsSpan(), out index)
                || int.TryParse(rawText, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
        }

        private static bool TryParseSharedStringIndex(ReadOnlySpan<char> text, out int index) {
            index = 0;
            if (text.Length == 0) {
                return false;
            }

            int parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return false;
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    return false;
                }

                parsed = (parsed * 10) + digit;
            }

            index = parsed;
            return true;
        }

        private object? ConvertByHints(ExcelCellValueType? type, uint? styleIndex, string? rawText, string? inlineText) =>
            ConvertByHints(type, styleIndex, rawText, inlineText, out _);

        private object? ConvertByHints(ExcelCellValueType? type, uint? styleIndex, string? rawText, string? inlineText, out bool customHandled) {
            customHandled = false;
            // Custom converter hook (cell-level). If provided and handled, honor it.
            var hook = _opt.CellValueConverter;
            if (hook != null) {
                var ctx = new ExcelCellContext(type, styleIndex, rawText, inlineText, _opt.Culture);
                var res = hook(ctx);
                if (res.Handled) { customHandled = true; return res.Value; }
                if (TryFormatDisplayText(type, styleIndex, rawText, out string displayed)) return displayed;
            }
            if (!string.IsNullOrEmpty(inlineText)) return inlineText;

            if (type == ExcelCellValueType.SharedString && TryParseSharedStringIndex(rawText, out var sstIndex))
                return GetSharedString(sstIndex);

            if (type == ExcelCellValueType.Boolean && rawText != null)
                return BoxBoolean(ExcelBooleanCellText.IsTrue(rawText));

            if (type == ExcelCellValueType.Number && rawText != null) {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && Styles.IsDateLike(styleIndex.Value)) {
                    if (TryParseInvariantDouble(rawText, out var oa)
                        || OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        return DateOrSerial(oa, styleIndex);
                }
                if (_opt.NumericAsDecimal) {
                    if (TryParseExcelNumberAsDecimal(rawText, _opt.Culture, out var dec))
                        return dec;
                    if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var dbl))
                        return dbl;
                    return rawText;
                } else {
                    if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var num))
                        return num;
                }
                return rawText;
            }

            if (type == ExcelCellValueType.Date && rawText != null) {
                if (ExcelIsoDateText.TryParse(rawText, _opt.Culture, System.Globalization.DateTimeStyles.AssumeLocal, out var dt))
                    return dt;
                return rawText;
            }

            if (type == ExcelCellValueType.String || type == ExcelCellValueType.InlineString || type == ExcelCellValueType.SharedString) {
                if (type == ExcelCellValueType.SharedString && TryParseSharedStringIndex(rawText, out var idx))
                    return GetSharedString(idx);
                return rawText != null ? Utilities.ExcelSanitizer.DecodeXString(rawText) : inlineText;
            }

            if (rawText != null) {
                if (_opt.TreatDatesUsingNumberFormat && styleIndex is not null && Styles.IsDateLike(styleIndex.Value)) {
                    if (TryParseInvariantDouble(rawText, out var oa)
                        || OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(rawText, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        return DateOrSerial(oa, styleIndex);
                    return rawText;
                }

                if (_opt.NumericAsDecimal) {
                    if (TryParseExcelNumberAsDecimal(rawText, _opt.Culture, out var dec2))
                        return dec2;
                    if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var dbl2))
                        return dbl2;
                } else {
                    if (TryParseExcelNumberAsDouble(rawText, _opt.Culture, out var num))
                        return num;
                }
                return rawText;
            }

            return null;
        }

        private struct CellRaw {
            public int Row;
            public int Col;
            public ExcelCellValueType? TypeHint;
            public uint? StyleIndex;
            public bool HasFormula;
            public bool CustomValueHandled;
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
            Unknown,
            Error
        }
    }
}
