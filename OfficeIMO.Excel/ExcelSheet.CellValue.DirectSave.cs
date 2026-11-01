using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        // Core implementation: single source of truth (no locks here)
        private void CellValueCore(int row, int column, object? value) {
            MaterializeDeferredDataSetImportIfNeeded();
            CellValueCoreNoMaterialize(row, column, value);
        }

        private bool TrySetPendingDirectCellValue(int row, int column, object? value)
            => TrySetPendingDirectCell(row, column, ExcelBufferedCell.FromValue(value));

        // Only worksheet settings whose XML the direct writers carry unchanged may precede
        // the first buffered cell; the rest depend on styles or parts those writers replace.
        private bool CanStartCellValueBuffer() {
            if (_worksheetPart.DrawingsPart != null || _worksheetPart.TableDefinitionParts.Any()) {
                return false;
            }

            foreach (var child in WorksheetRoot.ChildElements) {
                switch (child) {
                    case SheetData sheetData:
                        if (sheetData.Elements<Row>().Any(row => row.Elements<Cell>().Any())) {
                            return false;
                        }

                        break;
                    case Columns columns:
                        if (columns.Elements<Column>().Any(column => column.Style != null)) {
                            return false;
                        }

                        break;
                    case SheetViews:
                    case SheetFormatProperties:
                    case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting:
                    case SheetDimension:
                    case MergeCells:
                    case Hyperlinks:
                    case LegacyDrawing:
                    case AutoFilter:
                    case PageMargins:
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private bool TrySetPendingDirectCell(int row, int column, in ExcelBufferedCell cell) {
            if (HasCellValueWriteAccess) {
                return TrySetPendingDirectCellLocked(row, column, cell);
            }

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                return TrySetPendingDirectCellLocked(row, column, cell);
            } finally {
                lck.ExitWriteLock();
            }
        }

        private bool CanAcceptPendingDirectCell()
            => EnablePendingDirectCellValueBuffer
               && !_materializingPendingCellValueDirectSaveBuffer
               && !_disablePendingCellValueDirectSaveBuffer
               && !_excelDocument.HasDeferredDirectDataSetImport
               && !_excelDocument.HasDirectDataSetFastSaveState
               && (_pendingCellValueDirectSaveBuffer != null || (!_hasCellValueDomWrites && CanStartCellValueBuffer()));

        private bool TrySetPendingDirectCellLocked(int row, int column, in ExcelBufferedCell cell) {
            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            bool prepareDateTimeOffset = cell.Reference is DateTimeOffset;
            if (!CanAcceptPendingDirectCell()
                || (!prepareDateTimeOffset && !CanBufferDirectCell(cell))) {
                return false;
            }

            return TrySetPendingDirectCellCore(row, column, cell, prepareDateTimeOffset);
        }

        private bool TrySetPendingDirectCellFormula(int row, int column, string formula) {
            if (HasCellValueWriteAccess) {
                return TrySetPendingDirectCellFormulaLocked(row, column, formula);
            }

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                return TrySetPendingDirectCellFormulaLocked(row, column, formula);
            } finally {
                lck.ExitWriteLock();
            }
        }

        private bool TrySetPendingDirectCellFormulaLocked(int row, int column, string formula) {
            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (!EnablePendingDirectCellValueBuffer
                || _materializingPendingCellValueDirectSaveBuffer
                || _excelDocument.HasDeferredDirectDataSetImport
                || _excelDocument.HasDirectDataSetFastSaveState
                || (_pendingCellValueDirectSaveBuffer == null && _hasCellValueDomWrites)
                || (_pendingCellValueDirectSaveBuffer == null && !CanStartCellValueBuffer())) {
                return false;
            }

            object directValue = new DirectFormulaCellValue(Utilities.ExcelSanitizer.SanitizeFormula(formula));
            return TrySetPendingDirectCellCore(row, column, ExcelBufferedCell.FromValue(directValue));
        }

        private bool TrySetPendingDirectCellCore(int row, int column, in ExcelBufferedCell cell, bool prepareDateTimeOffset = false) {
            EnsureDynamicArrayCellWritable(row, column);
            if (!_excelDocument.TryReservePendingDirectCellValueSheet(this)) {
                return false;
            }

            int currentThreadId = Environment.CurrentManagedThreadId;
            if (_pendingCellValueDirectSaveBuffer == null) {
                _pendingCellValueDirectSaveThreadId = currentThreadId;
            } else if (_pendingCellValueDirectSaveThreadId != currentThreadId) {
                _disablePendingCellValueDirectSaveBuffer = true;
                MaterializePendingDirectCellValues();
                return false;
            }

            var buffer = _pendingCellValueDirectSaveBuffer ??= new ExcelCellValueBuffer();
            // Establish coordinate, sequence and thread admission before a
            // user conversion callback. A rejected buffer write uses the DOM
            // path, which must invoke that callback only once.
            bool appends = buffer.CanAdd(row, column);
            int existingCell = appends ? -1 : buffer.FindCell(row, column);
            if (!appends && (existingCell < 0 || _worksheetPart.TableDefinitionParts.Any())) {
                MaterializePendingDirectCellValues();
                return false;
            }

            ExcelBufferedCell bufferedCell = cell;
            if (prepareDateTimeOffset && cell.Reference is DateTimeOffset dateTimeOffset) {
                if (!TryPrepareDateTimeOffsetPendingDirectCellValue(dateTimeOffset, out DateTime converted)) {
                    MaterializePendingDirectCellValues();
                    CellDateTimeOffsetValueCore(row, column, dateTimeOffset, converted);
                    return true;
                }
                bufferedCell = ExcelBufferedCell.DateTimeCell(converted);
            }

            if (appends) {
                buffer.Add(row, column, bufferedCell);
                if (bufferedCell.HasLineBreak) {
                    SetBufferedAutomaticStyle(buffer, buffer.CellCount - 1, 0U, bufferedCell);
                }

                if (MirrorPendingDirectCellValueBufferToWorksheet || buffer.CellCount < PendingDirectCellValueMinimumCellCount) {
                    ApplyPendingDirectCellValueToDom(row, column, bufferedCell);
                }
            } else {
                ReplaceBufferedCell(buffer, existingCell, row, column, bufferedCell);
            }

            if (bufferedCell.Reference is not DirectFormulaCellValue) {
                _excelDocument.MarkFormulaInputMutation();
            }

            if (!_excelDocument.IsPackageDirty) {
                _excelDocument.MarkPackageDirty();
            }

            return true;
        }

        // Overwriting keeps the cell's format, as the DOM does.
        private void ReplaceBufferedCell(ExcelCellValueBuffer buffer, int cellIndex, int row, int column, in ExcelBufferedCell cell) {
            uint baseStyle = GetEffectiveBufferedStyle(buffer, cellIndex);
            buffer.Replace(cellIndex, cell);
            if (IsMirroredBufferedCell(cellIndex)) {
                ApplyPendingDirectCellValueToDom(row, column, cell);
                buffer.SetStyle(cellIndex, GetCell(row, column).StyleIndex?.Value ?? 0U);
                return;
            }

            buffer.SetStyle(cellIndex, baseStyle);
            SetBufferedAutomaticStyle(buffer, cellIndex, baseStyle, cell);
        }

        internal int MirroredCellCount => PendingDirectCellValueMinimumCellCount - 1;

        private bool IsMirroredBufferedCell(int cellIndex) => cellIndex < MirroredCellCount;

        // Date and time cells without a format take the automatic one when written.
        private uint GetEffectiveBufferedStyle(ExcelCellValueBuffer buffer, int cellIndex) {
            uint style = buffer.GetStyle(cellIndex);
            if (style != 0U) {
                return style;
            }

            ExcelBufferedCell cell = buffer.GetCell(cellIndex);
            if (cell.Kind == ExcelBufferedCellKind.DateTime || IsBufferedDateValue(cell.Reference)) {
                return GetAutomaticNumberFormatStyleIndex(0U, duration: false);
            }

            return IsBufferedTimeValue(cell.Reference)
                ? GetAutomaticNumberFormatStyleIndex(0U, duration: true)
                : 0U;
        }

        private static bool IsBufferedDateValue(object? value) {
#if NET6_0_OR_GREATER
            return value is DateOnly;
#else
            return false;
#endif
        }

        private static bool IsBufferedTimeValue(object? value) {
#if NET6_0_OR_GREATER
            return value is TimeSpan or TimeOnly;
#else
            return value is TimeSpan;
#endif
        }

        private void SetBufferedAutomaticStyle(ExcelCellValueBuffer buffer, int cellIndex, uint baseStyle, in ExcelBufferedCell cell) {
            object? value = cell.ToValue();
            var dataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cell.Kind == ExcelBufferedCellKind.Text
                ? DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                : DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
            if (RequiresAutomaticCellFormatting(value, dataType)) {
                Cell scratch = StyleScratchCell(baseStyle);
                ApplyAutomaticCellFormatting(scratch, value, dataType);
                buffer.SetStyle(cellIndex, scratch.StyleIndex?.Value ?? 0U);
            }
        }

        private Cell StyleScratchCell(uint styleIndex) {
            Cell scratch = _styleScratchCell ??= new Cell();
            scratch.StyleIndex = styleIndex == 0U ? null : styleIndex;
            return scratch;
        }

        private void WriteCellStyle(int row, int column, Action<Cell> apply) {
            bool buffered;
            if (HasCellValueWriteAccess) {
                buffered = TryWritePendingCellStyle(row, column, apply);
            } else {
                var lck = _excelDocument.EnsureLock();
                lck.EnterWriteLock();
                try {
                    buffered = TryWritePendingCellStyle(row, column, apply);
                } finally {
                    lck.ExitWriteLock();
                }
            }

            if (!buffered) {
                WriteLockConditional(() => apply(GetCell(row, column)));
            }
        }

        private bool TryWritePendingCellStyle(int row, int column, Action<Cell> apply) {
            if (!CanAcceptPendingDirectCell()) {
                return false;
            }

            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            int cellIndex = buffer?.FindCell(row, column) ?? -1;
            if (cellIndex < 0) {
                if (!TrySetPendingDirectCellCore(row, column, ExcelBufferedCell.StyleOnly)) {
                    return false;
                }

                buffer = _pendingCellValueDirectSaveBuffer!;
                cellIndex = buffer.CellCount - 1;
            }

            if (IsMirroredBufferedCell(cellIndex)) {
                Cell cell = GetCell(row, column);
                apply(cell);
                buffer!.SetStyle(cellIndex, cell.StyleIndex?.Value ?? 0U);
                return true;
            }

            Cell scratch = StyleScratchCell(GetEffectiveBufferedStyle(buffer!, cellIndex));
            apply(scratch);
            buffer!.SetStyle(cellIndex, scratch.StyleIndex?.Value ?? 0U);
            return true;
        }

        // A link cell is text; its relationship stays on the worksheet part.
        private bool TrySetPendingHyperlinkCell(int row, int column, string text) {
            if (!CanAcceptPendingDirectCell()) {
                return false;
            }

            ExcelBufferedCell cell = ExcelBufferedCell.Text(text);
            return CanBufferDirectCell(cell) && TrySetPendingDirectCellCore(row, column, cell);
        }

        private void ApplyPendingDirectCellValueToDom(int row, int column, in ExcelBufferedCell cell) {
            if (cell.Kind == ExcelBufferedCellKind.StyleOnly) {
                GetCell(row, column);
                return;
            }

            ApplyPendingDirectCellValueToDom(row, column, cell.ToValue());
        }

        private void ApplyPendingDirectCellValueToDom(int row, int column, object? value) {
            if (value is DirectFormulaCellValue formula) {
                CellFormulaCore(row, column, formula.Formula);
                return;
            }

            CellValueCoreNoMaterialize(row, column, value);
        }

        private static bool CanBufferDirectCell(in ExcelBufferedCell cell) {
            switch (cell.Kind) {
                case ExcelBufferedCellKind.Text:
                    CoerceValueHelper.ValidateSharedStringLength((string)cell.Reference!, "value");
                    return true;
                case ExcelBufferedCellKind.DateTime:
                    _ = cell.DateTimeValue.ToOADate();
                    return true;
                case ExcelBufferedCellKind.Value:
                    switch (cell.Reference) {
                        case float:
                        case decimal:
                        case short:
                        case uint:
                        case ulong:
                        case ushort:
                        case byte:
                        case sbyte:
                        case TimeSpan:
                        case DirectFormulaCellValue:
                            return true;
#if NET6_0_OR_GREATER
                        case DateOnly dateOnly:
                            _ = dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
                            return true;
                        case TimeOnly:
                            return true;
#endif
                        default:
                            return false;
                    }
                default:
                    return true;
            }
        }

        private bool TryPrepareDateTimeOffsetPendingDirectCellValue(DateTimeOffset value, out DateTime converted) {
            converted = ConvertDateTimeOffsetForCell(value);

            if (value.UtcDateTime < CellValueExcelMinimumSupportedDate) {
                return false;
            }

            try {
                _ = converted.ToOADate();
                return true;
            } catch (ArgumentException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        internal void MaterializePendingDirectCellValues(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null) {
                return;
            }

            _materializingPendingCellValueDirectSaveBuffer = true;
            try {
                // The first cells were mirrored as they arrived and may since have been edited in place.
                ApplyCellValueBufferToWorksheet(buffer, MirroredCellCount, cancellationToken);
                if (ReferenceEquals(_pendingCellValueDirectSaveBuffer, buffer)) {
                    _pendingCellValueDirectSaveBuffer = null;
                    _excelDocument.ClearPendingDirectCellValueSheet(this);
                }
            } finally {
                _materializingPendingCellValueDirectSaveBuffer = false;
            }
        }

        internal void ApplyCellValueBufferToWorksheet(ExcelCellValueBuffer buffer, int firstCellIndex, CancellationToken cancellationToken) {
            for (int rowIndex = 0; rowIndex < buffer.WrittenRowCount; rowIndex++) {
                int row = buffer.GetRowNumber(rowIndex);
                int end = buffer.GetEndCellIndex(rowIndex);
                for (int cellIndex = Math.Max(firstCellIndex, buffer.GetFirstCellIndex(rowIndex)); cellIndex < end; cellIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int column = buffer.GetColumn(cellIndex);
                    ApplyPendingDirectCellValueToDom(row, column, buffer.GetCell(cellIndex));
                    uint style = buffer.GetStyle(cellIndex);
                    if (style != 0U) {
                        GetCell(row, column).StyleIndex = style;
                    }
                }
            }
        }

        internal bool IsMaterializingPendingDirectCellValues
            => Volatile.Read(ref _materializingPendingCellValueDirectSaveBuffer);

        private bool CanWriteCellValueBuffer(ExcelCellValueBuffer buffer)
            => buffer.HasCompactRectangle && !WorksheetHoldsCellsOutsideBuffer(buffer);

        // Hands the buffer to an extended save; cells that cannot go direct stay for the DOM.
        internal ExcelCellValueBuffer? TryTakeCellValueBufferForDirectWorksheet() {
            var buffer = _pendingCellValueDirectSaveBuffer;
            return buffer != null && CanWriteCellValueBuffer(buffer) ? buffer : null;
        }

        internal void ReleaseCellValueBuffer() {
            _pendingCellValueDirectSaveBuffer = null;
            _excelDocument.ClearPendingDirectCellValueSheet(this);
        }

        internal ExcelDocument.DirectWorkbookCellStyles CreateWorkbookCellStyles()
            => new ExcelDocument.DirectWorkbookCellStyles(
                GetAutomaticNumberFormatStyleIndex(0U, duration: false),
                GetAutomaticNumberFormatStyleIndex(0U, duration: true));

        internal bool HasStyledPendingCells => _pendingCellValueDirectSaveBuffer?.HasCellStyles == true;

        internal int PendingCellValueCount => _pendingCellValueDirectSaveBuffer?.CellCount ?? 0;

        // A DOM-only cell inside the buffer's rectangle, such as an array formula set
        // between buffered cells, would be lost to a writer that sees only the buffer.
        private bool WorksheetHoldsCellsOutsideBuffer(ExcelCellValueBuffer buffer) {
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return false;
            }

            foreach (Row row in sheetData.Elements<Row>()) {
                foreach (Cell cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is not string reference) {
                        return true;
                    }

                    (int rowNumber, int column) = ExcelA1.ParseCellRef(reference);
                    if (rowNumber <= buffer.RowCount && column <= buffer.ColumnCount && !buffer.Contains(rowNumber, column)) {
                        return true;
                    }
                }
            }

            return false;
        }

        internal bool TryPromotePendingDirectCellValuesToSaveCandidate() {
            var buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null) {
                return false;
            }

            if (buffer.CellCount < PendingDirectCellValueMinimumCellCount
                || buffer.HasCellStyles
                || !CanWriteCellValueBuffer(buffer)
                || !IsOnlyVisibleWorksheetWithoutParts()) {
                return false;
            }

            bool registered = _excelDocument.RegisterDeferredCellValueBufferSaveCandidate(this, buffer);
            if (registered) {
                _pendingCellValueDirectSaveBuffer = null;
                _excelDocument.ClearPendingDirectCellValueSheet(this);
            }

            return registered;
        }
    }
}
