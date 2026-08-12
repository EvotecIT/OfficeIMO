namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Reuses record positions proved by the complete worksheet-validation pass when
    /// consecutive populated rows have one stable binary layout.
    /// </summary>
    internal sealed partial class XlsbTabularDataReader {
        private bool TryReadValidatedRowPlan(int rowIndex) {
            XlsbValidatedRowPlan? plan = _validatedRowPlan;
            if (plan == null || rowIndex < plan.FirstRowIndex || rowIndex > plan.LastRowIndex) {
                return false;
            }

            int start = _records.Position;
            long expectedStart = plan.FirstContentPosition
                + (long)(rowIndex - plan.FirstRowIndex) * plan.RowStride;
            if (start != expectedStart || start > _records.Length - plan.RowStride) {
                return false;
            }

            bool checkCancellation = _cancellationToken.CanBeCanceled;
            byte[] bytes = _records.Buffer;
            XlsbValidatedCellPlan[] cells = plan.Cells;
            for (int index = 0; index < cells.Length; index++) {
                if (checkCancellation) {
                    CheckCancellation();
                }

                XlsbValidatedCellPlan cell = cells[index];
                int payloadOffset = start + cell.RelativePayloadOffset;
                switch (cell.RecordType) {
                    case BrtCellRk:
                        StoreValidatedRkCell(bytes, payloadOffset);
                        break;
                    case BrtCellReal:
                        StoreValidatedRealCell(bytes, payloadOffset);
                        break;
                    case BrtCellIsst:
                        StoreValidatedSharedStringCell(bytes, payloadOffset);
                        break;
                    default:
                        StoreCellFast(bytes, cell.RecordType, payloadOffset);
                        break;
                }
            }

            _records.Position = start + plan.RowStride;
            _pendingRowIndex = checked(rowIndex + 1);
            _hasPendingRow = true;
            return true;
        }
    }

    internal sealed class XlsbValidatedRowPlanBuilder {
        private readonly List<XlsbValidatedCellPlan> _candidateCells = new List<XlsbValidatedCellPlan>();
        private XlsbValidatedCellPlan[]? _cells;
        private int _physicalRowIndex = -1;
        private int _currentRowIndex = -1;
        private int _currentContentPosition;
        private int _currentCellCount;
        private bool _currentMatches = true;
        private bool _disabled;
        private int _firstRowIndex;
        private int _lastRowIndex = -1;
        private int _firstContentPosition;
        private int _rowStride;
        private int _pendingCompletedRowIndex = -1;

        internal void BeginRow(int rowIndex, int contentPosition) {
            if (_physicalRowIndex >= 0) {
                CompleteCurrentRow(rowIndex, contentPosition);
            }

            _physicalRowIndex++;
            _currentRowIndex = rowIndex;
            _currentContentPosition = contentPosition;
            _currentCellCount = 0;
            _currentMatches = true;
            if (_cells == null) {
                _candidateCells.Clear();
            }
        }

        internal void ObserveCell(
            int recordType,
            int payloadPosition,
            int recordSize,
            int column) {
            if (_disabled || _physicalRowIndex <= 0) {
                return;
            }

            if (_currentCellCount == 0 && _pendingCompletedRowIndex >= 0) {
                _lastRowIndex = _pendingCompletedRowIndex;
                _pendingCompletedRowIndex = -1;
            }

            var cell = new XlsbValidatedCellPlan(
                recordType,
                payloadPosition - _currentContentPosition,
                recordSize,
                column);
            if (_cells == null) {
                _candidateCells.Add(cell);
            } else {
                if (_currentCellCount >= _cells.Length || !_cells[_currentCellCount].Equals(cell)) {
                    _currentMatches = false;
                }
            }

            _currentCellCount++;
        }

        internal XlsbValidatedRowPlan? Build() {
            if (_disabled || _cells == null || _lastRowIndex < _firstRowIndex) {
                return null;
            }

            return new XlsbValidatedRowPlan(
                _firstRowIndex,
                _lastRowIndex,
                _firstContentPosition,
                _rowStride,
                _cells);
        }

        private void CompleteCurrentRow(int nextRowIndex, int nextContentPosition) {
            if (_disabled || _physicalRowIndex == 0) {
                return;
            }

            int stride = nextContentPosition - _currentContentPosition;
            if (_currentCellCount == 0
                || nextRowIndex != _currentRowIndex + 1
                || stride <= 0) {
                _disabled = true;
                return;
            }

            if (_cells == null) {
                _cells = _candidateCells.ToArray();
                _firstRowIndex = _currentRowIndex;
                _firstContentPosition = _currentContentPosition;
                _rowStride = stride;
                _pendingCompletedRowIndex = _currentRowIndex;
                return;
            }

            if (!_currentMatches || _currentCellCount != _cells.Length || stride != _rowStride) {
                _disabled = true;
                return;
            }

            _pendingCompletedRowIndex = _currentRowIndex;
        }
    }

    internal sealed class XlsbValidatedRowPlan {
        internal XlsbValidatedRowPlan(
            int firstRowIndex,
            int lastRowIndex,
            int firstContentPosition,
            int rowStride,
            XlsbValidatedCellPlan[] cells) {
            FirstRowIndex = firstRowIndex;
            LastRowIndex = lastRowIndex;
            FirstContentPosition = firstContentPosition;
            RowStride = rowStride;
            Cells = cells;
        }

        internal int FirstRowIndex { get; }

        internal int LastRowIndex { get; }

        internal int FirstContentPosition { get; }

        internal int RowStride { get; }

        internal XlsbValidatedCellPlan[] Cells { get; }
    }

    internal readonly struct XlsbValidatedCellPlan : IEquatable<XlsbValidatedCellPlan> {
        internal XlsbValidatedCellPlan(
            int recordType,
            int relativePayloadOffset,
            int recordSize,
            int column) {
            RecordType = recordType;
            RelativePayloadOffset = relativePayloadOffset;
            RecordSize = recordSize;
            Column = column;
        }

        internal int RecordType { get; }

        internal int RelativePayloadOffset { get; }

        internal int RecordSize { get; }

        internal int Column { get; }

        public bool Equals(XlsbValidatedCellPlan other) =>
            RecordType == other.RecordType
            && RelativePayloadOffset == other.RelativePayloadOffset
            && RecordSize == other.RecordSize
            && Column == other.Column;

        public override bool Equals(object? obj) =>
            obj is XlsbValidatedCellPlan other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = RecordType;
                hash = (hash * 397) ^ RelativePayloadOffset;
                hash = (hash * 397) ^ RecordSize;
                return (hash * 397) ^ Column;
            }
        }
    }
}
