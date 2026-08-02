using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Tracks the aggregate populated-cell count across one workbook projection.</summary>
    internal sealed class XlsbCellReadBudget {
        private readonly int _maximum;
        private int _remaining;

        internal XlsbCellReadBudget(int maximum) {
            if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
            _maximum = maximum;
            _remaining = maximum;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Consume() {
            if (_remaining <= 0) {
                ThrowExceeded();
            }

            _remaining--;
        }

        internal void Consume(int count) {
            if (count < 0 || count > _remaining) {
                ThrowExceeded();
            }
            _remaining -= count;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowExceeded() =>
            throw new InvalidDataException(
                $"The XLSB workbook exceeds the configured limit of {_maximum} populated cells.");
    }
}
