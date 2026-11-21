namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private int _tableLayoutEpoch;

        /// <summary>Changes whenever a table range, header-row count or table part may have changed on any worksheet.</summary>
        internal int TableLayoutEpoch => Volatile.Read(ref _tableLayoutEpoch);

        /// <summary>Retires every worksheet's table-header index after table ranges, header counts or table parts change.</summary>
        internal void InvalidateTableLayouts() => Interlocked.Increment(ref _tableLayoutEpoch);
    }
}
