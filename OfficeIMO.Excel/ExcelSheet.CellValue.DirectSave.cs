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

        // A link cell is text plus a direct format; its relationship stays on the worksheet part.
        private bool TrySetPendingHyperlinkCell(int row, int column, string text, bool style) {
            ExcelBufferedCell cell = style ? ExcelBufferedCell.HyperlinkText(text) : ExcelBufferedCell.Text(text);
            if ((style && ExcelBufferedCell.Text(text).Style != ExcelBufferedCellStyle.None) || !CanAcceptPendingDirectCell()) {
                return false;
            }

            CoerceValueHelper.ValidateSharedStringLength(text, "value");
            var buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null || buffer.CanAdd(row, column)) {
                return TrySetPendingDirectCellCore(row, column, cell);
            }

            int cellIndex = buffer.FindCell(row, column);
            if (cellIndex < 0 || !HasDefaultCellFormat(buffer.GetCell(cellIndex))) {
                return false;
            }

            buffer.Replace(cellIndex, cell);
            if (cellIndex < PendingDirectCellValueMinimumCellCount - 1) {
                CellValueCoreNoMaterialize(row, column, text);
            }

            return true;
        }

        // Linking keeps a cell's format, which only the DOM can combine with the link's font.
        private static bool HasDefaultCellFormat(in ExcelBufferedCell cell)
            => cell.Style == ExcelBufferedCellStyle.None
               && cell.Kind != ExcelBufferedCellKind.DateTime
               && cell.Kind != ExcelBufferedCellKind.Value;

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
            if (!buffer.CanAdd(row, column)) {
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

            buffer.Add(row, column, bufferedCell);

            if (MirrorPendingDirectCellValueBufferToWorksheet || buffer.CellCount < PendingDirectCellValueMinimumCellCount) {
                ApplyPendingDirectCellValueToDom(row, column, bufferedCell.ToValue());
            } else if (bufferedCell.Reference is not DirectFormulaCellValue) {
                _excelDocument.MarkFormulaInputMutation();
            }

            if (!_excelDocument.IsPackageDirty) {
                _excelDocument.MarkPackageDirty();
            }

            return true;
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
                int mirroredCells = PendingDirectCellValueMinimumCellCount - 1;
                foreach (var cell in buffer.EnumerateWrittenCells()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (mirroredCells > 0) {
                        mirroredCells--;
                        continue;
                    }

                    ApplyPendingDirectCellValueToDom(cell.Row, cell.Column, cell.Value);
                }

                ApplyBufferedHyperlinkStyles(buffer);
                if (ReferenceEquals(_pendingCellValueDirectSaveBuffer, buffer)) {
                    _pendingCellValueDirectSaveBuffer = null;
                    _excelDocument.ClearPendingDirectCellValueSheet(this);
                }
            } finally {
                _materializingPendingCellValueDirectSaveBuffer = false;
            }
        }

        internal void ApplyCellValueBufferToWorksheet(ExcelCellValueBuffer buffer, CancellationToken cancellationToken) {
            foreach (var cell in buffer.EnumerateWrittenCells()) {
                cancellationToken.ThrowIfCancellationRequested();
                ApplyPendingDirectCellValueToDom(cell.Row, cell.Column, cell.Value);
            }

            ApplyBufferedHyperlinkStyles(buffer);
        }

        private void ApplyBufferedHyperlinkStyles(ExcelCellValueBuffer buffer) {
            foreach ((int row, int column) in buffer.EnumerateHyperlinkCells()) {
                ApplyHyperlinkStyle(GetCell(row, column));
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

        internal bool HasStyledPendingCells => _pendingCellValueDirectSaveBuffer is { } buffer
            && (buffer.UsesStyle(ExcelBufferedCellStyle.WrapText) || buffer.UsesStyle(ExcelBufferedCellStyle.Hyperlink));

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
