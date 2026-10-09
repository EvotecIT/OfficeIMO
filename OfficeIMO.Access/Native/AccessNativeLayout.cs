namespace OfficeIMO.Access {
    /// <summary>Physical generation coordinates used by the shared bounded native reader.</summary>
    internal sealed class AccessNativeLayout {
        internal AccessNativeLayout(bool jet3) { IsJet3 = jet3; }
        internal bool IsJet3 { get; }
        internal int PageSize => IsJet3 ? 2048 : 4096;
        internal int DataRowCount => IsJet3 ? 8 : 12;
        internal int DataRowDirectory => DataRowCount + 2;
        internal int RowColumnCountSize => IsJet3 ? 1 : 2;
        internal int TableHeaderSize => IsJet3 ? 43 : 63;
        internal int TableRowCount => IsJet3 ? 12 : 16;
        internal int TableMaxColumns => IsJet3 ? 21 : 41;
        internal int TableMaxVariableColumns => TableMaxColumns + 2;
        internal int TableColumnCount => IsJet3 ? 25 : 45;
        internal int TableLogicalIndexes => TableColumnCount + 2;
        internal int TablePhysicalIndexes => TableColumnCount + 6;
        internal int TableOwnedPages => IsJet3 ? 35 : 55;
        internal int TableFreePages => TableOwnedPages + 4;
        internal int IndexStatisticsSize => IsJet3 ? 8 : 12;
        internal int ColumnSize => IsJet3 ? 18 : 25;
        internal int ColumnNumber => IsJet3 ? 1 : 5;
        internal int ColumnVariableIndex => ColumnNumber + 2;
        internal int ColumnFlags => IsJet3 ? 13 : 15;
        internal int ColumnFixedOffset => IsJet3 ? 14 : 21;
        internal int ColumnLength => IsJet3 ? 16 : 23;
        internal int ColumnSortOrder => IsJet3 ? 9 : 11;
        internal int PhysicalIndexSize => IsJet3 ? 39 : 52;
        internal int PhysicalIndexColumns => IsJet3 ? 0 : 4;
        internal int PhysicalIndexOwnedPages => PhysicalIndexColumns + 30;
        internal int PhysicalIndexRoot => IsJet3 ? 34 : 38;
        internal int PhysicalIndexFlags => IsJet3 ? 38 : 46;
        internal int LogicalIndexSize => IsJet3 ? 20 : 28;
        internal int LogicalIndexNumber => IsJet3 ? 0 : 4;
        internal int LogicalIndexPhysical => LogicalIndexNumber + 4;
        internal int LogicalIndexRelated => IsJet3 ? 9 : 13;
        internal int LogicalIndexTable => LogicalIndexRelated + 4;
        internal int LogicalIndexCascade => IsJet3 ? 17 : 21;
        internal int LogicalIndexType => IsJet3 ? 19 : 23;
    }
}
