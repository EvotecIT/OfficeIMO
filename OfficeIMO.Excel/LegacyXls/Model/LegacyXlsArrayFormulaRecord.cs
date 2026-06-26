namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an Array BIFF record discovered while importing a legacy XLS worksheet.
    /// </summary>
    public sealed class LegacyXlsArrayFormulaRecord {
        internal LegacyXlsArrayFormulaRecord(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            ushort optionFlags,
            int formulaTokenByteCount,
            int formulaExtraByteCount,
            int matchedFormulaCellCount,
            bool formulaTextProjected,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            if (firstRow <= 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (firstColumn <= 0) throw new ArgumentOutOfRangeException(nameof(firstColumn));
            if (lastRow < firstRow) throw new ArgumentOutOfRangeException(nameof(lastRow));
            if (lastColumn < firstColumn) throw new ArgumentOutOfRangeException(nameof(lastColumn));
            if (formulaTokenByteCount < 0) throw new ArgumentOutOfRangeException(nameof(formulaTokenByteCount));
            if (formulaExtraByteCount < 0) throw new ArgumentOutOfRangeException(nameof(formulaExtraByteCount));
            if (matchedFormulaCellCount < 0) throw new ArgumentOutOfRangeException(nameof(matchedFormulaCellCount));
            if (payloadLength < 0) throw new ArgumentOutOfRangeException(nameof(payloadLength));

            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            OptionFlags = optionFlags;
            FormulaTokenByteCount = formulaTokenByteCount;
            FormulaExtraByteCount = formulaExtraByteCount;
            MatchedFormulaCellCount = matchedFormulaCellCount;
            FormulaTextProjected = formulaTextProjected;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the first one-based row covered by the array formula.</summary>
        public int FirstRow { get; }

        /// <summary>Gets the first one-based column covered by the array formula.</summary>
        public int FirstColumn { get; }

        /// <summary>Gets the last one-based row covered by the array formula.</summary>
        public int LastRow { get; }

        /// <summary>Gets the last one-based column covered by the array formula.</summary>
        public int LastColumn { get; }

        /// <summary>Gets the A1 range covered by the array formula.</summary>
        public string Range {
            get {
                string start = A1.CellReference(FirstRow, FirstColumn);
                string end = A1.CellReference(LastRow, LastColumn);
                return string.Equals(start, end, StringComparison.Ordinal) ? start : start + ":" + end;
            }
        }

        /// <summary>Gets the raw Array record option flags.</summary>
        public ushort OptionFlags { get; }

        /// <summary>Gets whether the array formula asks Excel to calculate during the next recalculation.</summary>
        public bool AlwaysCalculate => (OptionFlags & 0x0001) != 0;

        /// <summary>Gets the byte count of the parsed formula token stream.</summary>
        public int FormulaTokenByteCount { get; }

        /// <summary>Gets the byte count of ancillary parsed-formula data after the token stream.</summary>
        public int FormulaExtraByteCount { get; }

        /// <summary>Gets the number of cached formula cells matched to this Array record during import.</summary>
        public int MatchedFormulaCellCount { get; }

        /// <summary>Gets the number of cells covered by the declared array formula range.</summary>
        public int DeclaredCellCount => checked((LastRow - FirstRow + 1) * (LastColumn - FirstColumn + 1));

        /// <summary>Gets whether OfficeIMO projected formula text onto at least one matched cached formula cell.</summary>
        public bool FormulaTextProjected { get; }

        /// <summary>Gets the byte offset of the Array BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF payload length in bytes.</summary>
        public int PayloadLength { get; }
    }
}
