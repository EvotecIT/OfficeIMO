namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Tracks the aggregate populated-cell count across one workbook projection.</summary>
    internal sealed class XlsbCellReadBudget {
        private readonly int _maximum;
        private int _count;

        internal XlsbCellReadBudget(int maximum) {
            if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
            _maximum = maximum;
        }

        internal void Consume() {
            _count = checked(_count + 1);
            if (_count > _maximum) {
                throw new InvalidDataException(
                    $"The XLSB workbook exceeds the configured limit of {_maximum} populated cells.");
            }
        }
    }
}
