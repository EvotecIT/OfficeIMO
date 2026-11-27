using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// A cell read from the worksheet XML the workbook loaded, with the changes recorded since applied. Its writes are
    /// recorded as changes too, so the owned save writes them and a later SDK opening replays them.
    /// </summary>
    internal sealed class ExcelXmlFormulaCellRecord : ExcelFormulaCellRecord {
        private readonly ExcelXmlFormulaSheetModel _model;

        internal ExcelXmlFormulaCellRecord(ExcelXmlFormulaSheetModel model, int row, int column) {
            _model = model;
            Row = row;
            Column = column;
        }

        internal readonly int Row;
        internal readonly int Column;
        // A canonical reference is made again from the position when asked for; any other is kept as written.
        internal string? WrittenReference;
        internal bool CanonicalReference;
        internal uint? Style;
        internal ExcelCellValueType? Type;
        internal string? Value;
        internal bool Inline;
        internal string InlineValue = string.Empty;
        internal uint? ValueMeta;
        internal uint? CellMeta;
        internal bool CellMetaPresent;
        internal bool Formula;
        internal string? FormulaValue;
        internal ExcelFormulaKind Kind;
        internal uint? Shared;
        internal string? FormulaRange;
        internal bool? Calculate;
        internal bool Children;

        internal override object Identity => this;
        internal override bool HasFormula => Formula;
        internal override string? FormulaText => FormulaValue;
        internal override ExcelFormulaKind FormulaKind => Kind;
        internal override uint? SharedIndex => Shared;
        internal override string? FormulaReference => FormulaRange;
        internal override bool? CalculateCell => Calculate;
        internal override string? Reference => CanonicalReference ? WrittenReference ??= ExcelA1.CellReference(Row, Column) : WrittenReference;
        internal override string? ValueText => Value;
        internal override ExcelCellValueType? ValueType => Type;
        internal override uint? ValueMetaIndex => ValueMeta;
        internal override uint? CellMetaIndex => CellMeta;
        internal override bool HasCellMetaIndex => CellMetaPresent;
        internal override uint? StyleIndex => Style;
        internal override bool HasInlineString => Inline;
        internal override string InlineText => Inline ? InlineValue : string.Empty;
        internal override bool HasChildren => Children || Formula || Value != null || Inline;

        internal override bool TryGetPosition(out int row, out int column) {
            row = Row;
            column = Column;
            return true;
        }

        // A workbook read from its XML has no cell metadata, so it has no rich values.
        internal override string? ResolveRichValueError(string? fallback) => fallback;

        internal override void SetCalculateCell(bool? calculate) {
            if (Calculate == calculate) return;
            Calculate = calculate;
            _model.Record(this);
        }

        internal override void SetValue(string? text, ExcelCellValueType? type) {
            if (string.Equals(Value, text, StringComparison.Ordinal) && Type == type) return;
            Value = text;
            Type = type;
            _model.Record(this);
        }

        internal override void ClearInlineString() {
            if (!Inline) return;
            Inline = false;
            InlineValue = string.Empty;
            _model.Record(this);
        }

        internal override void ClearValueMetaIndex() {
            if (ValueMeta == null) return;
            ValueMeta = null;
            _model.Record(this);
        }

        internal override void SetFormulaReference(string reference) =>
            _model.InModel(this, "Formula range change").SetFormulaReference(reference);

        internal override void SetOwnFormula(string text) =>
            _model.InModel(this, "Shared formula materialization").SetOwnFormula(text);
    }

    /// <summary>
    /// A worksheet read with the owned XML reader from the package the workbook loaded, with the workbook's recorded
    /// changes applied in order. It reads the part once; later changes apply on top as they are recorded.
    /// </summary>
    internal sealed class ExcelXmlFormulaSheetModel : ExcelFormulaSheetModel {
        private readonly ExcelDocument _document;
        private readonly ExcelSheet _sheet;
        private readonly ExcelDeferredSheet _entry;
        private readonly ExcelXmlFormulaWorkbookModel _workbook;
        private Dictionary<long, ExcelXmlFormulaCellRecord> _cells = new();
        private readonly List<ExcelXmlFormulaCellRecord> _order = new();
        private readonly List<ExcelXmlFormulaCellRecord> _formulas = new();
        private readonly List<string> _merged = new();
        private Dictionary<ExcelFormulaCell, (int Row, int Column)>? _coordinates;
        private List<ExcelFormulaTable>? _tables;
        private readonly OfficeCharBuffer _text = new(256);
        private ExcelDeferredCellStore? _appliedStore;
        private int _applied;

        private ExcelXmlFormulaSheetModel(ExcelDocument document, ExcelSheet sheet, ExcelDeferredSheet entry, ExcelXmlFormulaWorkbookModel workbook, byte[] packageBytes) {
            _document = document;
            _sheet = sheet;
            _entry = entry;
            _workbook = workbook;
            PackageBytes = packageBytes;
            PartUri = new Uri(entry.PartName, UriKind.Relative);
        }

        /// <summary>The package bytes the sheet was read from; a save or an SDK opening replaces them.</summary>
        internal byte[] PackageBytes { get; }

        internal override Uri PartUri { get; }

        internal bool SheetRequestsFullCalculation { get; private set; }

        internal override bool HasPivotTables => _workbook.HasPivotTables(_entry);

        /// <summary>Reads the worksheet, or returns null when its markup needs the model.</summary>
        internal static ExcelXmlFormulaSheetModel? TryRead(ExcelDocument document, ExcelSheet sheet, ExcelDeferredSheet entry, ExcelXmlFormulaWorkbookModel workbook) {
            var model = new ExcelXmlFormulaSheetModel(document, sheet, entry, workbook, workbook.PackageBytes);
            return model.TryParse() ? model : null;
        }

        private static long Key(int row, int column) => ((long)row << 15) | (uint)column;

        /// <summary>Applies the changes recorded since the last call; false when the sheet must be read again.</summary>
        internal bool TrySync() {
            ExcelDeferredCellStore? store = _entry.Cells;
            if (store == null) {
                // The recorded changes went into a save, which replaced the package bytes, or were never made.
                return _appliedStore == null;
            }
            if (!ReferenceEquals(store, _appliedStore)) {
                if (_appliedStore != null) return false;
                _appliedStore = store;
                _applied = 0;
            }
            if (store.Count < _applied) return false;
            for (; _applied < store.Count; _applied++) Apply(store[_applied]);
            return true;
        }

        private void Apply(in ExcelDeferredCellEdit edit) {
            if (edit.Kind == ExcelDeferredCellKind.Style) return;
            ExcelXmlFormulaCellRecord cell = GetOrCreateRecord(edit.Row, edit.Column);
            cell.Inline = false;
            cell.InlineValue = string.Empty;
            switch (edit.Kind) {
                case ExcelDeferredCellKind.Text:
                    cell.Type = ExcelCellValueType.InlineString;
                    cell.Value = null;
                    cell.Inline = true;
                    cell.InlineValue = edit.Text;
                    break;
                case ExcelDeferredCellKind.Number:
                    cell.Type = ExcelCellValueType.Number;
                    cell.Value = edit.Number.ToString("R", CultureInfo.InvariantCulture);
                    break;
                case ExcelDeferredCellKind.Integer:
                    cell.Type = ExcelCellValueType.Number;
                    cell.Value = ((long)edit.Number).ToString(CultureInfo.InvariantCulture);
                    break;
                case ExcelDeferredCellKind.Decimal:
                    cell.Type = ExcelCellValueType.Number;
                    cell.Value = edit.Text;
                    break;
                case ExcelDeferredCellKind.Boolean:
                    cell.Type = ExcelCellValueType.Boolean;
                    cell.Value = edit.Number != 0D ? "1" : "0";
                    break;
                case ExcelDeferredCellKind.Date:
                    cell.Type = ExcelCellValueType.Number;
                    cell.Value = ExcelDateSystemConverter.ToSerial(edit.Date, _document.DateSystem).ToString("R", CultureInfo.InvariantCulture);
                    break;
                case ExcelDeferredCellKind.FormulaResult:
                    cell.Type = edit.FormulaResultType;
                    cell.Value = edit.FormulaResultText;
                    cell.Calculate = edit.FormulaResultCalculate;
                    break;
            }
        }

        private ExcelXmlFormulaCellRecord GetOrCreateRecord(int row, int column) {
            long key = Key(row, column);
            if (_cells.TryGetValue(key, out ExcelXmlFormulaCellRecord? cell)) return cell;
            cell = new ExcelXmlFormulaCellRecord(this, row, column) { CanonicalReference = true };
            _cells.Add(key, cell);
            _order.Add(cell);
            _coordinates = null;
            return cell;
        }

        /// <summary>Records a write for the owned save, or sends it to the model when the workbook has opened one.</summary>
        internal void Record(ExcelXmlFormulaCellRecord cell) {
            if (_document.RecordDeferredFormulaResult(_sheet, _entry, cell)) {
                // The change is in the store already applied to this record.
                if (ReferenceEquals(_entry.Cells, _appliedStore) || _appliedStore == null) {
                    _appliedStore = _entry.Cells;
                    _applied = _entry.Cells!.Count;
                }
                return;
            }
            ExcelFormulaCellRecord target = InModel(cell, null);
            target.SetValue(cell.Value, cell.Type);
            target.SetCalculateCell(cell.Calculate);
            if (!cell.Inline) target.ClearInlineString();
            if (cell.ValueMeta == null) target.ClearValueMetaIndex();
        }

        /// <summary>The same cell in the model the workbook opens for a change its XML cannot record.</summary>
        internal ExcelFormulaCellRecord InModel(ExcelXmlFormulaCellRecord cell, string? reason) {
            if (reason != null) _document.OpenDeferredSdk(reason + " on worksheet '" + _entry.Name + "'");
            ExcelSheet sheet = _document.FindSheetForFormulas(_entry.Name)
                ?? throw new InvalidOperationException("The worksheet '" + _entry.Name + "' is no longer part of the workbook.");
            return sheet.FormulaModel.GetOrCreate(cell.Row, cell.Column);
        }

        internal override ExcelFormulaCellRecord? Find(int row, int column) =>
            _cells.TryGetValue(Key(row, column), out ExcelXmlFormulaCellRecord? cell) ? cell : null;

        internal override ExcelFormulaCellRecord GetOrCreate(int row, int column) => GetOrCreateRecord(row, column);

        internal override void RemoveIfEmpty(ExcelFormulaCellRecord cell) {
            // A spilled value is never cleared here: dynamic arrays need cell metadata, which this model never reads.
        }

        internal override IEnumerable<ExcelFormulaCellRecord> FormulaCells() {
            foreach (ExcelXmlFormulaCellRecord cell in _formulas) yield return cell;
        }

        internal override IEnumerable<ExcelFormulaCellRecord> Cells() {
            foreach (ExcelXmlFormulaCellRecord cell in _order) yield return cell;
        }

        internal override IReadOnlyDictionary<ExcelFormulaCell, (int Row, int Column)> EffectiveCoordinates() {
            Dictionary<ExcelFormulaCell, (int Row, int Column)>? coordinates = _coordinates;
            if (coordinates != null) return coordinates;
            coordinates = new Dictionary<ExcelFormulaCell, (int Row, int Column)>(_order.Count);
            foreach (ExcelXmlFormulaCellRecord cell in _order) coordinates[new ExcelFormulaCell(cell)] = (cell.Row, cell.Column);
            return _coordinates = coordinates;
        }

        internal override IEnumerable<string> MergedRanges() => _merged;

        internal override IReadOnlyList<ExcelFormulaTable> Tables() => _tables ??= _workbook.ReadTables(_entry);

        internal override bool HasTableHeaders => Tables().Count != 0;

        internal override bool IsTableHeader(int top, int left, int bottom, int right) => ExcelFormulaTables.HeaderIntersects(Tables(), top, left, bottom, right);

        internal override ExcelCellData ReadValue(int row, int column) =>
            ExcelFormulaValues.Read(Find(row, column), _workbook.SharedString);

        internal override bool TryReadPlainValue(int row, int column, out double? number, out string? text) {
            number = null;
            text = null;
            if (!_cells.TryGetValue(Key(row, column), out ExcelXmlFormulaCellRecord? cell) || cell.Formula || cell.Inline) return false;
            if (cell.Type is null or ExcelCellValueType.Number) {
                if (!OfficeRoundTripText.TryParse(cell.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) return false;
                number = parsed;
                text = cell.Value;
                return true;
            }
            if (cell.Type != ExcelCellValueType.SharedString || string.IsNullOrEmpty(cell.Value)
                || !int.TryParse(cell.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int index)) return false;
            text = _workbook.SharedString(index) ?? string.Empty;
            return true;
        }

        private bool TryParse() {
            OfficePooledBytes content;
            try {
                using var source = new MemoryStream(PackageBytes, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                if (!package.PartExists(_entry.PartName)) return false;
                OfficeXmlPullReader reader = package.OpenXml(_entry.PartName, out content, processMarkupCompatibility: false);
                using (content) return Parse(reader, content.Length);
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or IOException
                or FormatException or OverflowException) {
                return false;
            }
        }

        private bool Parse(OfficeXmlPullReader reader, int length) {
            int previousRow = 0;
            bool root = false;
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.Depth == 1) {
                    if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "worksheet")) return false;
                    root = true;
                    continue;
                }
                if (reader.Depth != 2 || reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                if (reader.LocalNameEquals("dimension")) {
                    // The dimension sizes the cell index; a sparse sheet's wide dimension is capped by its size.
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string dimension)
                        && ExcelA1.TryParseRange(dimension, out int top, out int left, out int bottom, out int right)) {
                        long cells = Math.Min((long)(bottom - top + 1) * (right - left + 1), Math.Min(length / 24, 1 << 20));
                        if (cells > 16 && _cells.Count == 0) {
                            _cells = new Dictionary<long, ExcelXmlFormulaCellRecord>((int)cells);
                            _order.Capacity = (int)cells;
                        }
                    }
                } else if (reader.LocalNameEquals("sheetCalcPr")) {
                    SheetRequestsFullCalculation = reader.TryGetAttribute(OfficeXmlNamespaces.None, "fullCalcOnLoad", out string full)
                        && full is "1" or "true";
                } else if (reader.LocalNameEquals("sheetData")) {
                    if (reader.IsEmptyElement) continue;
                    while (reader.Read()) {
                        if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 2) break;
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 3) continue;
                        if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "row")) {
                            reader.SkipElement();
                            continue;
                        }
                        int rowNumber = previousRow + 1;
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string rowText)
                            && uint.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsedRow) && parsedRow > 0U) {
                            if (parsedRow > ExcelA1.MaxRows) return false;
                            rowNumber = (int)parsedRow;
                        }
                        previousRow = rowNumber;
                        if (rowNumber > ExcelA1.MaxRows) return false;
                        if (!reader.IsEmptyElement && !ParseRow(reader, rowNumber)) return false;
                    }
                } else if (reader.LocalNameEquals("mergeCells")) {
                    if (reader.IsEmptyElement) continue;
                    while (reader.Read()) {
                        if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 2) break;
                        if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 3
                            && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "mergeCell")) {
                            _merged.Add(reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string merged) ? merged : string.Empty);
                        }
                    }
                }
            }
            return root;
        }

        private bool ParseRow(OfficeXmlPullReader reader, int rowNumber) {
            int previousColumn = 0;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 3) return true;
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 4) continue;
                if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "c")) {
                    reader.SkipElement();
                    continue;
                }
                int column = previousColumn + 1;
                int row = rowNumber;
                string? reference = null;
                bool canonical = false;
                uint? style = null, cellMeta = null, valueMeta = null;
                ExcelCellValueType? type = null;
                bool cellMetaPresent = false;
                byte[] data = reader.Data;
                // The attributes are read from their bytes; only a reference that is not canonical becomes a string.
                for (int index = 0; index < reader.AttributeCount; index++) {
                    if (reader.GetAttributeNamespace(index, out _) != OfficeXmlNamespaces.None) continue;
                    reader.GetAttributeRawValue(index, out int start, out int length);
                    switch (reader.GetAttributeLocalName(index)) {
                        case "r":
                            if (TryParseCanonicalReference(data, start, length, out int explicitRow, out int explicitColumn)) {
                                row = explicitRow;
                                column = explicitColumn;
                                canonical = true;
                            } else {
                                reference = reader.GetAttributeValue(index);
                                if (ExcelA1.TryParseCellReferenceFast(reference, out explicitRow, out explicitColumn)) {
                                    row = explicitRow;
                                    column = explicitColumn;
                                }
                            }
                            break;
                        case "s":
                            if (TryParseUInt(reader, index, start, length, out uint parsedStyle)) style = parsedStyle;
                            break;
                        case "t":
                            type = Array.IndexOf(data, (byte)'&', start, length) >= 0 ? ParseType(reader.GetAttributeValue(index)) : ParseType(data, start, length);
                            break;
                        case "cm":
                            cellMetaPresent = true;
                            if (TryParseUInt(reader, index, start, length, out uint parsedCellMeta)) cellMeta = parsedCellMeta;
                            break;
                        case "vm":
                            if (TryParseUInt(reader, index, start, length, out uint parsedValueMeta)) valueMeta = parsedValueMeta;
                            break;
                    }
                }
                previousColumn = column;
                if (column > ExcelA1.MaxColumns || row > ExcelA1.MaxRows) return false;
                // A later cell at the same position replaces an earlier one, as a lookup by reference finds the first.
                long key = Key(row, column);
                if (_cells.ContainsKey(key)) {
                    reader.SkipElement();
                    continue;
                }
                var cell = new ExcelXmlFormulaCellRecord(this, row, column) {
                    WrittenReference = reference, CanonicalReference = canonical, Style = style, Type = type,
                    CellMeta = cellMeta, CellMetaPresent = cellMetaPresent, ValueMeta = valueMeta
                };
                if (!reader.IsEmptyElement) ParseCellChildren(reader, cell);
                _cells.Add(key, cell);
                _order.Add(cell);
                if (cell.Formula) _formulas.Add(cell);
            }
            return false;
        }

        private void ParseCellChildren(OfficeXmlPullReader reader, ExcelXmlFormulaCellRecord cell) {
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == 4) return;
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 5) continue;
                if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "f")) {
                    if (cell.Formula) {
                        reader.SkipElement();
                        continue;
                    }
                    cell.Formula = true;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "t", out string kind)) {
                        cell.Kind = kind switch {
                            "shared" => ExcelFormulaKind.Shared,
                            "array" => ExcelFormulaKind.Array,
                            "dataTable" => ExcelFormulaKind.DataTable,
                            _ => ExcelFormulaKind.Normal,
                        };
                    }
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "si", out string shared)
                        && uint.TryParse(shared, NumberStyles.None, CultureInfo.InvariantCulture, out uint sharedIndex)) cell.Shared = sharedIndex;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string range)) cell.FormulaRange = range;
                    if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "ca", out string calculate)) cell.Calculate = ParseBoolean(calculate);
                    cell.FormulaValue = ReadElementText(reader);
                } else if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "v")) {
                    string text = ReadElementText(reader);
                    if (cell.Value == null) cell.Value = text;
                } else if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "is")) {
                    cell.Inline = true;
                    cell.InlineValue = ReadInlineText(reader);
                } else {
                    cell.Children = true;
                    reader.SkipElement();
                }
            }
        }

        private string ReadElementText(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            _text.Clear();
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (reader.NodeKind == OfficeXmlNodeKind.Text) reader.AppendText(_text);
            }
            return _text.Length == 0 ? string.Empty : _text.ToString();
        }

        // The cell's own text element, or else its runs' text, as cell reads take it.
        private string ReadInlineText(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            string? direct = null;
            StringBuilder runs = new StringBuilder();
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.Depth == depth + 1 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "t")) {
                    string text = ReadElementText(reader);
                    direct ??= text;
                } else if (reader.Depth == depth + 2 && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "t")) {
                    runs.Append(Utilities.ExcelSanitizer.DecodeXString(ReadElementText(reader)));
                } else if (reader.Depth == depth + 1 && !reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "r")) {
                    reader.SkipElement();
                }
            }
            return direct != null ? Utilities.ExcelSanitizer.DecodeXString(direct) : runs.ToString();
        }

        // A value written with character references is decoded first.
        private static bool TryParseUInt(OfficeXmlPullReader reader, int index, int start, int length, out uint value) =>
            TryParseUInt(reader.Data, start, length, out value)
            || Array.IndexOf(reader.Data, (byte)'&', start, length) >= 0
                && uint.TryParse(reader.GetAttributeValue(index), NumberStyles.None, CultureInfo.InvariantCulture, out value);

        private static bool TryParseUInt(byte[] data, int start, int length, out uint value) {
            value = 0U;
            if (length == 0 || length > 10) return false;
            ulong parsed = 0UL;
            for (int index = start; index < start + length; index++) {
                int digit = data[index] - (byte)'0';
                if ((uint)digit > 9U) return false;
                parsed = parsed * 10UL + (uint)digit;
            }
            if (parsed > uint.MaxValue) return false;
            value = (uint)parsed;
            return true;
        }

        // Upper-case column letters, then a row number without a leading zero.
        private static bool TryParseCanonicalReference(byte[] data, int start, int length, out int row, out int column) {
            row = column = 0;
            int index = start, end = start + length;
            while (index < end && data[index] >= (byte)'A' && data[index] <= (byte)'Z' && index - start < 3) {
                column = column * 26 + data[index] - (byte)'A' + 1;
                index++;
            }
            if (index == start || index == end || data[index] == (byte)'0' || end - index > 7) return false;
            for (; index < end; index++) {
                int digit = data[index] - (byte)'0';
                if ((uint)digit > 9U) return false;
                row = row * 10 + digit;
            }
            return row <= ExcelA1.MaxRows && column <= ExcelA1.MaxColumns;
        }

        private static ExcelCellValueType ParseType(byte[] data, int start, int length) {
            if (length == 1) {
                switch (data[start]) {
                    case (byte)'b': return ExcelCellValueType.Boolean;
                    case (byte)'n': return ExcelCellValueType.Number;
                    case (byte)'e': return ExcelCellValueType.Error;
                    case (byte)'s': return ExcelCellValueType.SharedString;
                    case (byte)'d': return ExcelCellValueType.Date;
                }
            } else if (length == 9 && Encoding.ASCII.GetString(data, start, length) == "inlineStr") {
                return ExcelCellValueType.InlineString;
            }
            return ExcelCellValueType.String;
        }

        private static ExcelCellValueType ParseType(string type) => type switch {
            "b" => ExcelCellValueType.Boolean,
            "n" => ExcelCellValueType.Number,
            "e" => ExcelCellValueType.Error,
            "s" => ExcelCellValueType.SharedString,
            "inlineStr" => ExcelCellValueType.InlineString,
            "d" => ExcelCellValueType.Date,
            _ => ExcelCellValueType.String,
        };

        private static bool? ParseBoolean(string value) => value switch {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null,
        };
    }

    /// <summary>The workbook-level parts formulas read from the package the workbook loaded.</summary>
    internal sealed class ExcelXmlFormulaWorkbookModel : ExcelFormulaWorkbookModel {
        private const string TableRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "table";
        private const string PivotTableRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "pivotTable";
        private const string SharedStringsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "sharedStrings";
        private readonly ExcelDocument _document;
        private readonly ExcelDeferredWorkbook _workbook;
        private readonly object _gate = new();
        private List<ExcelFormulaDefinedName>? _definedNames;
        private bool _fullCalculation;
        private SharedStringCache? _sharedStrings;
        private bool _workbookRead;

        internal ExcelXmlFormulaWorkbookModel(ExcelDocument document, ExcelDeferredWorkbook workbook, byte[] packageBytes) {
            _document = document;
            _workbook = workbook;
            PackageBytes = packageBytes;
        }

        internal byte[] PackageBytes { get; }

        internal ExcelDeferredWorkbook Workbook => _workbook;

        internal override bool HasCellMetadata => false;

        internal override ExcelCellMetadataIndex? CellMetadata => null;

        internal override List<ExcelFormulaSheetEntry> SheetEntries() {
            var entries = new List<ExcelFormulaSheetEntry>(_workbook.Sheets.Count);
            foreach (ExcelDeferredSheet sheet in _workbook.Sheets) entries.Add(new ExcelFormulaSheetEntry(sheet.Name, sheet.IsWorksheet, sheet));
            return entries;
        }

        internal override List<ExcelFormulaDefinedName> DefinedNames() {
            ReadWorkbook();
            return new List<ExcelFormulaDefinedName>(_definedNames!);
        }

        internal override ExcelSheet CreateSheet(ExcelFormulaSheetEntry entry) => new ExcelSheet(_document, (ExcelDeferredSheet)entry.Element);

        /// <summary>Reads the workbook part and the shared strings; false when either needs the model.</summary>
        internal bool TryReadParts() {
            try {
                ReadWorkbook();
                SharedStringCache strings = ReadSharedStrings();
                _ = strings.Count;
                _sharedStrings = strings;
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or IOException
                or FormatException or OverflowException or XlsxTabularFastPathNotSupportedException or InvalidOperationException) {
                return false;
            }
        }

        internal override IEnumerable<(ExcelFormulaSheetEntry Sheet, ExcelFormulaTable Table)> Tables() {
            foreach (ExcelDeferredSheet sheet in _workbook.Sheets) {
                if (!sheet.IsWorksheet) continue;
                var entry = new ExcelFormulaSheetEntry(sheet.Name, true, sheet);
                foreach (ExcelFormulaTable table in ReadTables(sheet)) yield return (entry, table);
            }
        }

        internal override bool RequestsFullCalculation(ExcelSheet sheet) {
            ReadWorkbook();
            return _fullCalculation || sheet.FormulaModel is ExcelXmlFormulaSheetModel { SheetRequestsFullCalculation: true };
        }

        internal string? SharedString(int index) {
            SharedStringCache? strings = _sharedStrings;
            if (strings == null) {
                lock (_gate) strings = _sharedStrings ??= ReadSharedStrings();
            }
            return strings.Get(index);
        }

        private SharedStringCache ReadSharedStrings() {
            using var source = new MemoryStream(PackageBytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            string? part = package.GetRelatedPartName(_workbook.WorkbookPartName, SharedStringsRelationshipType);
            if (part == null || !package.PartExists(part)) return SharedStringCache.Empty(new ExcelReaderOptions());
            OfficePooledBytes bytes = package.ReadPart(part);
            byte[] copy;
            using (bytes) {
                copy = new byte[bytes.Length];
                Buffer.BlockCopy(bytes.Array, 0, copy, 0, bytes.Length);
            }
            return SharedStringCache.BuildKeepingUnmarkedWhitespace(() => new MemoryStream(copy, writable: false), new ExcelReaderOptions());
        }

        internal bool HasPivotTables(ExcelDeferredSheet sheet) {
            using var source = new MemoryStream(PackageBytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(sheet.PartName)) {
                if (string.Equals(relationship.Type, PivotTableRelationshipType, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private void ReadWorkbook() {
            if (_workbookRead) return;
            lock (_gate) {
                if (_workbookRead) return;
                var names = new List<ExcelFormulaDefinedName>();
                using (var source = new MemoryStream(PackageBytes, writable: false))
                using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                    OfficeXmlPullReader reader = package.OpenXml(_workbook.WorkbookPartName, out OfficePooledBytes content, processMarkupCompatibility: false);
                    using (content) {
                        bool inNames = false, namesSeen = false;
                        while (reader.Read()) {
                            if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                            if (reader.Depth == 2) {
                                inNames = false;
                                if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "definedNames") && !namesSeen) {
                                    inNames = namesSeen = true;
                                } else if (reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "calcPr")) {
                                    _fullCalculation = IsTrue(reader, "fullCalcOnLoad") || IsTrue(reader, "forceFullCalc");
                                }
                                continue;
                            }
                            if (reader.Depth != 3 || !inNames || !reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "definedName")) continue;
                            string? name = reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string nameText) ? nameText : null;
                            uint? localSheetId = reader.TryGetAttribute(OfficeXmlNamespaces.None, "localSheetId", out string local)
                                && uint.TryParse(local, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsedLocal) ? parsedLocal : null;
                            string text = ReadText(reader);
                            names.Add(new ExcelFormulaDefinedName(name, text, localSheetId));
                        }
                    }
                }
                _definedNames = names;
                _workbookRead = true;
            }
        }

        private static bool IsTrue(OfficeXmlPullReader reader, string attribute) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, attribute, out string value) && value is "1" or "true";

        private static string ReadText(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            var builder = new StringBuilder();
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (reader.NodeKind == OfficeXmlNodeKind.Text) builder.Append(reader.GetText());
            }
            return builder.ToString();
        }

        /// <summary>The tables of a worksheet, read from its table parts.</summary>
        internal List<ExcelFormulaTable> ReadTables(ExcelDeferredSheet sheet) {
            var tables = new List<ExcelFormulaTable>();
            using var source = new MemoryStream(PackageBytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(sheet.PartName)) {
                if (!string.Equals(relationship.Type, TableRelationshipType, StringComparison.Ordinal)
                    || relationship.IsExternal || relationship.TargetPartName is not { } target || !package.PartExists(target)) continue;
                OfficeXmlPullReader reader = package.OpenXml(target, out OfficePooledBytes content, processMarkupCompatibility: false);
                using (content) {
                    string? name = null, displayName = null, reference = null;
                    uint headerRowCount = 1U, totalsRowCount = 1U;
                    bool totalsRowShown = false, root = false, inColumns = false;
                    var columns = new List<string?>();
                    while (reader.Read()) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (reader.Depth == 1) {
                            if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "table")) break;
                            root = true;
                            name = reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string n) ? n : null;
                            displayName = reader.TryGetAttribute(OfficeXmlNamespaces.None, "displayName", out string d) ? d : null;
                            reference = reader.TryGetAttribute(OfficeXmlNamespaces.None, "ref", out string r) ? r : null;
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "headerRowCount", out string h)
                                && uint.TryParse(h, NumberStyles.None, CultureInfo.InvariantCulture, out uint header)) headerRowCount = header;
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "totalsRowCount", out string t)
                                && uint.TryParse(t, NumberStyles.None, CultureInfo.InvariantCulture, out uint totals)) totalsRowCount = totals;
                            totalsRowShown = IsTrue(reader, "totalsRowShown");
                            continue;
                        }
                        if (reader.Depth == 2) {
                            inColumns = reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "tableColumns");
                            continue;
                        }
                        if (reader.Depth == 3 && inColumns && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "tableColumn")) {
                            columns.Add(reader.TryGetAttribute(OfficeXmlNamespaces.None, "name", out string column) ? column : null);
                        }
                    }
                    if (root) tables.Add(new ExcelFormulaTable(name, displayName, reference, headerRowCount, totalsRowCount, totalsRowShown, columns));
                }
            }
            return tables;
        }
    }

    /// <summary>Table helpers shared by the formula models.</summary>
    internal static class ExcelFormulaTables {
        /// <summary>Whether the range covers a cell of the header row of a table that has one.</summary>
        internal static bool HeaderIntersects(IReadOnlyList<ExcelFormulaTable> tables, int top, int left, int bottom, int right) {
            foreach (ExcelFormulaTable table in tables) {
                if (table.HeaderRowCount == 0U
                    || !ExcelA1.TryParseRange(table.Reference ?? string.Empty, out int headerRow, out int firstColumn, out _, out int lastColumn)) continue;
                if (headerRow >= top && headerRow <= bottom && firstColumn <= right && lastColumn >= left) return true;
            }
            return false;
        }
    }

    /// <summary>Cell values as cell reads give them, from any model's cell record.</summary>
    internal static class ExcelFormulaValues {
        internal static ExcelCellData Read(ExcelFormulaCellRecord? cell, Func<int, string?> sharedString) {
            if (cell == null) return new ExcelCellData(ExcelCellDataKind.Blank, null);
            string? cached = cell.ResolveRichValueError(cell.ValueText);
            ExcelCellValueType? type = cell.ValueType;
            if (cell.HasFormula) {
                object? formulaValue = cached;
                if (cached != null) {
                    if (type == ExcelCellValueType.Boolean) {
                        formulaValue = cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase);
                    } else if (type is ExcelCellValueType.String or ExcelCellValueType.SharedString or ExcelCellValueType.InlineString) {
                        formulaValue = Text(cell, sharedString);
                    } else if (type != ExcelCellValueType.Error
                        && OfficeRoundTripText.TryParse(cached, NumberStyles.Float, CultureInfo.InvariantCulture, out double cachedNumber)) {
                        formulaValue = cachedNumber;
                    }
                }
                return new ExcelCellData(ExcelCellDataKind.Formula, formulaValue, cell.FormulaText, cached);
            }

            if (cell.ValueText == null && !cell.HasInlineString) return new ExcelCellData(ExcelCellDataKind.Blank, null);
            if (type == ExcelCellValueType.Boolean) {
                return new ExcelCellData(ExcelCellDataKind.Boolean, cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase), cachedText: cached);
            }
            if (type == ExcelCellValueType.Error) return new ExcelCellData(ExcelCellDataKind.Error, cached, cachedText: cached);

            string text = Text(cell, sharedString);
            if (type is ExcelCellValueType.String or ExcelCellValueType.SharedString or ExcelCellValueType.InlineString) {
                return new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
            }
            if (OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new ExcelCellData(ExcelCellDataKind.Number, number, cachedText: text);
            }
            return string.IsNullOrEmpty(text)
                ? new ExcelCellData(ExcelCellDataKind.Blank, null)
                : new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
        }

        /// <summary>The cell's text: its shared string, its inline string, or its value as written.</summary>
        internal static string Text(ExcelFormulaCellRecord cell, Func<int, string?> sharedString) {
            ExcelCellValueType? type = cell.ValueType;
            if (type == ExcelCellValueType.SharedString) {
                string? raw = cell.ValueText;
                return !string.IsNullOrEmpty(raw) && int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
                    ? sharedString(index) ?? string.Empty : string.Empty;
            }
            if (type == ExcelCellValueType.InlineString) return cell.InlineText;
            string text = cell.ResolveRichValueError(cell.ValueText) ?? string.Empty;
            return type == ExcelCellValueType.String ? Utilities.ExcelSanitizer.DecodeXString(text) : text;
        }
    }
}
