namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single cell-level difference between two worksheet ranges.
    /// </summary>
    public sealed class ExcelRangeDifference {
        /// <summary>
        /// Creates a range difference.
        /// </summary>
        public ExcelRangeDifference(
            ExcelRangeDifferenceKind kind,
            int leftRow,
            int leftColumn,
            int rightRow,
            int rightColumn,
            string leftCell,
            string rightCell,
            object? leftValue,
            object? rightValue) {
            Kind = kind;
            LeftRow = leftRow;
            LeftColumn = leftColumn;
            RightRow = rightRow;
            RightColumn = rightColumn;
            LeftCell = leftCell;
            RightCell = rightCell;
            LeftValue = leftValue;
            RightValue = rightValue;
        }

        /// <summary>Difference kind.</summary>
        public ExcelRangeDifferenceKind Kind { get; }

        /// <summary>1-based row in the left range.</summary>
        public int LeftRow { get; }

        /// <summary>1-based column in the left range.</summary>
        public int LeftColumn { get; }

        /// <summary>1-based row in the right range.</summary>
        public int RightRow { get; }

        /// <summary>1-based column in the right range.</summary>
        public int RightColumn { get; }

        /// <summary>A1 cell reference in the left worksheet.</summary>
        public string LeftCell { get; }

        /// <summary>A1 cell reference in the right worksheet.</summary>
        public string RightCell { get; }

        /// <summary>Value from the left worksheet.</summary>
        public object? LeftValue { get; }

        /// <summary>Value from the right worksheet.</summary>
        public object? RightValue { get; }
    }
}
