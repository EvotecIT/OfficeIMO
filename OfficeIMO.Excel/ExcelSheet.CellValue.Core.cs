using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private readonly struct CellValueSharedStringIndexCacheEntry {
            internal CellValueSharedStringIndexCacheEntry(int index, bool containsLineBreak) {
                Index = index;
                ContainsLineBreak = containsLineBreak;
            }

            internal int Index { get; }

            internal bool ContainsLineBreak { get; }
        }

        /// <summary>The error Excel shows for a number it cannot represent, used for NaN and infinities.</summary>
        internal const string NonFiniteNumberErrorText = "#NUM!";

        private void CellValueCoreNoMaterialize(int row, int column, object? value) {
            // Materialization replays physical cells already admitted to the buffer.
            if (!_materializingPendingCellValueDirectSaveBuffer && !_excelDocument.IsMaterializingDeferredDataSetImport
                && TrySetTableHeaderValue(row, column, value)) return;
            if (value == null || value == DBNull.Value) {
                CellEmptyStringValueCore(row, column);
                return;
            }

            switch (value) {
                case string text:
                    CellStringValueCore(row, column, text);
                    return;
                case double number:
                    CellDoubleValueCore(row, column, number);
                    return;
                case float number:
                    CellDoubleValueCore(row, column, (double)number);
                    return;
                case decimal number:
                    CellDecimalValueCore(row, column, number);
                    return;
                case int number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case long number:
                    CellIntegerValueCore(row, column, InvariantNumberText.Get(number), ExcelExactNumber.FitsNumberCell(number));
                    return;
                case short number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case uint number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case ulong number:
                    CellIntegerValueCore(row, column, InvariantNumberText.Get(number), ExcelExactNumber.FitsNumberCell(number));
                    return;
                case ushort number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case byte number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case sbyte number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case bool boolean:
                    CellBooleanValueCore(row, column, boolean);
                    return;
                case DateTime dateTime:
                    CellDateTimeValueCore(row, column, dateTime);
                    return;
                case DateTimeOffset dateTimeOffset:
                    CellDateTimeOffsetValueCore(row, column, dateTimeOffset);
                    return;
#if NET6_0_OR_GREATER
                case DateOnly dateOnly:
                    CellDateOnlyValueCore(row, column, dateOnly);
                    return;
                case TimeOnly timeOnly:
                    CellTimeOnlyValueCore(row, column, timeOnly);
                    return;
#endif
                case TimeSpan timeSpan:
                    CellTimeSpanValueCore(row, column, timeSpan);
                    return;
            }

            var (cellValue, dataType) = CoerceForCell(value);

            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = cellValue;
            cell.DataType = dataType;
            ApplyAutomaticCellFormatting(cell, value, dataType);
            CompleteCellValueMutation(row, column);
        }

        private void CellStringValueCore(int row, int column, string? value) {
            if (string.IsNullOrEmpty(value)) {
                CellEmptyStringValueCore(row, column);
                return;
            }

            CoerceValueHelper.ValidateSharedStringLength(value!, nameof(value));
            if (TryStoreText(row, column, value!)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            string text = value!;
            if (TryGetCellValueSharedStringIndex(text, out int cachedSharedStringIndex, out bool cachedContainsLineBreak)) {
                SetExistingCellSharedStringValue(cell, cachedSharedStringIndex, cachedContainsLineBreak);
            } else if (_excelDocument.TryGetOrAddSharedStringIndexBelowLimit(
                    text,
                    CellValuePlainStringPromotionSharedStringCount,
                    validateNewString: true,
                    out int sharedStringIndex,
                    out bool containsLineBreak)) {
                AddCellValueSharedStringIndex(text, sharedStringIndex, containsLineBreak);
                SetExistingCellSharedStringValue(cell, sharedStringIndex, containsLineBreak);
            } else {
                SetExistingCellPlainStringValue(cell, text);
            }

            CompleteCellValueMutation(row, column);
        }

        private bool TryGetCellValueSharedStringIndex(string text, out int index, out bool containsLineBreak) {
            SharedStringTable? table = _excelDocument.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable;
            if (_cellValueSharedStringIndexCache != null && (_cellValueSharedStringTable == null
                || !_cellValueSharedStringTable.TryGetTarget(out SharedStringTable? cachedTable)
                || !ReferenceEquals(cachedTable, table))) {
                _cellValueSharedStringIndexCache = null;
            }
            if (_cellValueSharedStringIndexCache != null
                && _cellValueSharedStringIndexCache.TryGetValue(text, out CellValueSharedStringIndexCacheEntry entry)) {
                index = entry.Index;
                containsLineBreak = entry.ContainsLineBreak;
                return true;
            }

            index = -1;
            containsLineBreak = false;
            return false;
        }

        private void AddCellValueSharedStringIndex(string text, int index, bool containsLineBreak) {
            SharedStringTable? table = _excelDocument.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable;
            if (table == null) return;
            if (_cellValueSharedStringTable == null) _cellValueSharedStringTable = new WeakReference<SharedStringTable>(table);
            else _cellValueSharedStringTable.SetTarget(table);
            var cache = _cellValueSharedStringIndexCache;
            if (cache == null) {
                cache = new Dictionary<string, CellValueSharedStringIndexCacheEntry>(StringComparer.Ordinal);
                _cellValueSharedStringIndexCache = cache;
            } else if (cache.Count >= CellValueSharedStringIndexCacheLimit) {
                return;
            }

            cache[text] = new CellValueSharedStringIndexCacheEntry(index, containsLineBreak);
        }

        private void CellEmptyStringValueCore(int row, int column) {
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            ExcelWorksheetScan.SetValueText(cell, string.Empty);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            ExcelWorksheetScan.RemoveInlineString(cell);
            CompleteCellValueMutation(row, column);
        }

        private void SetExistingCellSharedStringValue(Cell cell, string value, int sharedStringIndex) {
            SetExistingCellSharedStringValue(cell, sharedStringIndex, value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0);
        }

        private void SetExistingCellSharedStringValue(Cell cell, int sharedStringIndex, bool containsLineBreak) {
            ExcelWorksheetScan.SetValueText(cell, SharedStringIndexText.Get(sharedStringIndex));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            ExcelWorksheetScan.RemoveInlineString(cell);
            if (containsLineBreak) {
                ApplyWrapText(cell);
            }
        }

        private void SetExistingCellPlainStringValue(Cell cell, string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            ExcelWorksheetScan.SetValueText(cell, Utilities.ExcelSanitizer.EncodeXString(value));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            ExcelWorksheetScan.RemoveInlineString(cell);
            if (value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0) {
                ApplyWrapText(cell);
            }
        }

        private void CellDoubleValueCore(int row, int column, double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                // Excel has no NaN or infinite cell value; it shows #NUM! for such a result.
                var errorCell = GetWritableValueCell(row, column);
                ClearCellValueMetadata(errorCell);
                ExcelWorksheetScan.SetValueText(errorCell, NonFiniteNumberErrorText);
                errorCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
                ExcelWorksheetScan.RemoveInlineString(errorCell);
                CompleteCellValueMutation(row, column);
                return;
            }

            string text = InvariantNumberText.Get(value);
            if (TryStoreNumberText(row, column, text)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellDecimalValueCore(int row, int column, decimal value) {
            string text = value.ToString(CultureInfo.InvariantCulture);
            if (!ExcelExactNumber.FitsNumberCell(value)) {
                // Excel keeps 15 significant digits; text keeps the exact value.
                CellStringValueCore(row, column, text);
                return;
            }

            if (TryStoreNumberText(row, column, text)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellIntegerValueCore(int row, int column, string text, bool fitsNumberCell) {
            if (fitsNumberCell) CellNumberTextValueCore(row, column, text);
            else CellStringValueCore(row, column, text);
        }

        private void CellNumberTextValueCore(int row, int column, string text) {
            if (TryStoreNumberText(row, column, text)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellBooleanValueCore(int row, int column, bool value) {
            if (TryStoreBoolean(row, column, value)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            ExcelWorksheetScan.SetValueText(cell, value ? "1" : "0");
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
            CompleteCellValueMutation(row, column);
        }

        private void CellDateTimeValueCore(int row, int column, DateTime value) {
            double serial = ExcelDateSystemConverter.ToSerial(value, _excelDocument.DateSystem);
            string text = InvariantNumberText.Get(serial);
            if (TryStoreTemporal(row, column, text, duration: false)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);
            CompleteCellValueMutation(row, column);
        }

        private DateTime ConvertDateTimeOffsetForCell(DateTimeOffset value) {
            try {
                return _excelDocument.DateTimeOffsetWriteStrategy(value);
            } catch (Exception ex) {
                throw new InvalidOperationException("The configured DateTimeOffset write strategy threw an exception.", ex);
            }
        }

        private bool WritesDateTimeOffsetAsText => _excelDocument.DateTimeOffsetWriteMode == ExcelDateTimeOffsetWriteMode.Text;

        private void CellDateTimeOffsetValueCore(int row, int column, DateTimeOffset value) {
            if (WritesDateTimeOffsetAsText) {
                CellStringValueCore(row, column, CoerceValueHelper.FormatDateTimeOffsetText(value));
                return;
            }

            CellDateTimeOffsetValueCore(row, column, value, ConvertDateTimeOffsetForCell(value));
        }

        private void CellDateTimeOffsetValueCore(int row, int column, DateTimeOffset value, DateTime converted) {
            string? numericText = null;
            if (value.UtcDateTime >= CellValueExcelMinimumSupportedDate) {
                try {
                    double serial = ExcelDateSystemConverter.ToSerial(converted, _excelDocument.DateSystem);
                    numericText = InvariantNumberText.Get(serial);
                } catch (ArgumentException) {
                    // Fall back to ISO text below for values Excel cannot represent numerically.
                } catch (OverflowException) {
                    // Fall back to ISO text below for values Excel cannot represent numerically.
                }
            }

            if (numericText != null) {
                var numericCell = GetWritableValueCell(row, column);
                ClearCellValueMetadata(numericCell);
                numericCell.CellValue = new CellValue(numericText);
                numericCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;

                uint baseStyleIndex = numericCell.StyleIndex?.Value ?? 0U;
                numericCell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);

                CompleteCellValueMutation(row, column);
                return;
            }

            string fallbackText = CoerceValueHelper.FormatDateTimeOffsetText(value);
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(fallbackText, validateNewString: true, out bool containsLineBreak);
            var fallbackCell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(fallbackCell);
            SetExistingCellSharedStringValue(fallbackCell, sharedStringIndex, containsLineBreak);
            CompleteCellValueMutation(row, column);
        }

#if NET6_0_OR_GREATER
        private void CellDateOnlyValueCore(int row, int column, DateOnly value) {
            string text = InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(
                value.ToDateTime(TimeOnly.MinValue),
                _excelDocument.DateSystem));
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);
            CompleteCellValueMutation(row, column);
        }

        private void CellTimeOnlyValueCore(int row, int column, TimeOnly value) {
            string text = InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(value.ToTimeSpan()));
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: true);
            CompleteCellValueMutation(row, column);
        }
#endif

        private void CellFormulaCore(int row, int column, string formula) {
            var safe = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            Cell cell = GetWritableValueCell(row, column);
            bool retainsCachedValue = ExcelWorksheetScan.FindValue(cell) != null;
            ClearCellValueMetadata(cell);
            // Excel formulas in XML should not start with '=' and must not include illegal control characters
            var cellFormula = new CellFormula(safe);
            if (cell.FirstChild == null) cell.AppendChild(cellFormula);
            else cell.CellFormula = cellFormula;
            if (retainsCachedValue) {
                cellFormula.CalculateCell = true;
            }
            _excelDocument.MarkFormulaAuthored(
                _worksheetPart,
                ExcelA1.CellReference(row, column),
                retainedCachedValue: retainsCachedValue);
            ClearHeaderCacheForCellMutation(row, column);
        }

        private void CellTimeSpanValueCore(int row, int column, TimeSpan value) {
            double serial = ExcelDateSystemConverter.ToSerial(value);
            string text = InvariantNumberText.Get(serial);
            if (TryStoreTemporal(row, column, text, duration: true)) return;
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            ExcelWorksheetScan.SetValueText(cell, text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: true);
            CompleteCellValueMutation(row, column);
        }

        private void CompleteCellValueMutation(int row, int column) {
            _excelDocument.MarkFormulaInputMutation();
            ClearHeaderCacheForCellMutation(row, column);
        }

        // Core coercion logic shared between sequential and parallel operations
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCell(object? value) {
            var dateTimeOffsetStrategy = _excelDocument.EffectiveDateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    int idx = _excelDocument.GetSharedStringIndex(s);
                    return new CellValue(SharedStringIndexText.Get(idx));
                },
                dateTimeOffsetStrategy,
                _excelDocument.DateSystem);
            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }
    }
}

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies deferred defined-name formula marks. It runs from package access, which may already hold the
        /// document lock, so it marks the cells without taking it.
        /// </summary>
        /// <summary>
        /// Defers formula marks owed by defined-name edits around a literal value write. A sheet with tables, or one whose
        /// worksheet is not loaded, keeps the immediate marks: a header write can regenerate totals formulas.
        /// </summary>
        private ExcelDocument.DefinedNameFormulaMarkDeferral DeferDefinedNameFormulaMarksForValueWrite() {
            if (!_excelDocument.HasPendingDefinedNameFormulaMarks) return default;
            // Read through the unguarded field: the guarded one would apply the marks this defers.
            WorksheetPart part = BoundWorksheetPart;
            if (!part.IsRootElementLoaded || part.Worksheet is not Worksheet worksheet
                || worksheet.GetFirstChild<TableParts>() != null) return default;
            return _excelDocument.DeferDefinedNameFormulaMarks();
        }

        /// <summary>Writes an imported formula with an explicit text result using the OOXML formula-string representation.</summary>
        internal void CellFormulaWithTextCache(int row, int column,
            string formula, string cachedText) {
            ArgumentNullException.ThrowIfNull(formula);
            ArgumentNullException.ThrowIfNull(cachedText);
            CoerceValueHelper.ValidateSharedStringLength(cachedText, nameof(cachedText));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            WriteLock(() => {
                MaterializePendingDirectCellValues();
                MaterializeDeferredDataSetImportIfNeeded();
                Cell cell = GetWritableValueCell(row, column);
                ClearCellValueMetadata(cell);
                SetExistingCellPlainStringValue(cell, cachedText);
                CompleteCellValueMutation(row, column);
                CellFormulaCore(row, column, safeFormula);
            });
        }

        /// <summary>
        /// Sets a shared-free array formula over a range. The top-left cell owns the formula metadata.
        /// </summary>
        public void SetArrayFormula(string a1Range, string formula) => SetArrayFormulaCore(a1Range, formula, dynamic: false);

        /// <summary>
        /// Authors a dynamic array at one anchor cell. Calculation determines its bounded spill range.
        /// </summary>
        public void SetDynamicArrayFormula(string a1Cell, string formula) {
            var (row, column) = ExcelA1.ParseCellRef(a1Cell);
            if (row < 1 || row > ExcelA1.MaxRows || column < 1 || column > ExcelA1.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(a1Cell));
            string anchor = ExcelA1.CellReference(row, column);
            SetArrayFormulaCore(anchor + ":" + anchor, formula, dynamic: true);
        }

        private void SetArrayFormulaCore(string a1Range, string formula, bool dynamic) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula, qualifyArrayFunctions: true);
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);
            WriteLock(() => {
                EnsureTableHeaderFormulaRangeWritable(r1, c1, r2, c2);
                foreach (var cell in ExcelWorksheetScan.FormulaCells(WorksheetRoot).Where(c => ExcelWorksheetScan.FindFormula(c)?.FormulaType?.Value == CellFormulaValues.Array).ToList()) {
                    string? reference = ExcelWorksheetScan.FindFormula(cell)?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive((r1, c1, r2, c2), (existingR1, existingC1, existingR2, existingC2))) {
                        throw new InvalidOperationException($"Array formula range '{a1Range}' overlaps existing array formula range '{reference}'.");
                    }
                }

                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = ExcelWorksheetScan.FindValue(topLeft) != null;
                ClearCellValueMetadata(topLeft);
                topLeft.CellMetaIndex = dynamic ? EnsureDynamicArrayMetadata() : null;
                topLeft.CellFormula = new CellFormula(safeFormula) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    ExcelWorksheetScan.FindFormula(topLeft)!.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    ExcelA1.CellReference(r1, c1),
                    retainedCachedValue: retainsCachedValue);
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        if (row == r1 && column == c1) continue;
                        var cell = GetCell(row, column);
                        ClearCellValueMetadata(cell);
                        cell.CellFormula = null;
                        cell.CellValue = null;
                    }
                }

                InvalidateDynamicArrayWriteIndex();
            });
        }

        internal void SetLegacyArrayFormula(string a1Range, string formula, System.Threading.CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            string safeFormula = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            int r2;
            int c2;
            if (!ExcelA1.TryParseRange(a1Range, out int r1, out int c1, out r2, out c2)) {
                (r1, c1) = ExcelA1.ParseCellRef(a1Range);
                if (r1 <= 0 || c1 <= 0) {
                    throw new ArgumentException($"Invalid A1 range '{a1Range}'.", nameof(a1Range));
                }
                r2 = r1;
                c2 = c1;
            }

            WriteLock(() => {
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var member = GetCell(row, column);
                        ClearCellValueMetadata(member);
                        if (row != r1 || column != c1) member.CellFormula = null;
                    }
                }
                var topLeft = GetCell(r1, c1);
                bool retainsCachedValue = ExcelWorksheetScan.FindValue(topLeft) != null;
                topLeft.CellFormula = new CellFormula(safeFormula) {
                    FormulaType = CellFormulaValues.Array,
                    Reference = a1Range
                };
                if (retainsCachedValue) {
                    ExcelWorksheetScan.FindFormula(topLeft)!.CalculateCell = true;
                }
                _excelDocument.MarkFormulaAuthored(
                    _worksheetPart,
                    ExcelA1.CellReference(r1, c1),
                    retainedCachedValue: retainsCachedValue);
            });
        }

        /// <summary>
        /// Clears an array formula whose reference overlaps the supplied range or cell.
        /// </summary>
        public void ClearArrayFormula(string a1RangeOrCell) {
            var bounds = a1RangeOrCell.IndexOf(':') >= 0
                ? ExcelA1.ParseRange(a1RangeOrCell)
                : CellAsRange(a1RangeOrCell);

            WriteLock(() => {
                var arrayCells = ExcelWorksheetScan.FormulaCells(WorksheetRoot)
                    .Where(c => ExcelWorksheetScan.FindFormula(c)?.FormulaType?.Value == CellFormulaValues.Array).ToList();
                foreach (Cell array in arrayCells) {
                    if (TryFixedArrayBounds(ExcelWorksheetScan.FindFormula(array)?.Reference?.Value ?? string.Empty,
                        out int top, out int left, out int bottom, out int right)
                        && RangesOverlapInclusive(bounds, (top, left, bottom, right)))
                        EnsureTableHeaderValuesNotCleared(top, left, bottom, right);
                }
                foreach (var cell in arrayCells) {
                    string? reference = ExcelWorksheetScan.FindFormula(cell)?.Reference?.Value;
                    if (!string.IsNullOrWhiteSpace(reference)
                        && TryFixedArrayBounds(reference!, out int existingR1, out int existingC1, out int existingR2, out int existingC2)
                        && RangesOverlapInclusive(bounds, (existingR1, existingC1, existingR2, existingC2))) {
                        foreach (Cell spillCell in ExcelWorksheetScan.Cells(WorksheetRoot).ToList()) {
                            if (!TryParseCellReference(spillCell.CellReference?.Value ?? "", out int row, out int column)
                                || row < existingR1 || row > existingR2 || column < existingC1 || column > existingC2)
                                continue;
                            spillCell.CellFormula = null;
                            spillCell.CellValue = null;
                            ClearCellValueMetadataAttribute(spillCell);
                            spillCell.DataType = null;
                            if (ReferenceEquals(spillCell, cell)) spillCell.CellMetaIndex = null;
                            DynamicSpillCacheSnapshots.Remove(DynamicCellKey(row, column));
                        }
                        SpillOwnership.WrittenOwners.Remove(DynamicCellKey(existingR1, existingC1));
                    }
                }

                InvalidateDynamicArrayWriteIndex();
            });
        }
    }
}
