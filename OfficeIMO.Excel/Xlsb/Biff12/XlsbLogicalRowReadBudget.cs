using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Tracks logical rows emitted across one XLSB workbook data-reader operation.</summary>
    internal sealed class XlsbLogicalRowReadBudget {
        private readonly int _maximum;
        private int _remaining;

        internal XlsbLogicalRowReadBudget(int maximum) {
            if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
            _maximum = maximum;
            _remaining = maximum;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Consume() {
            if (_remaining <= 0) ThrowExceeded();
            _remaining--;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowExceeded() =>
            throw new InvalidDataException(
                $"The XLSB workbook exceeds the configured limit of {_maximum} logical data rows.");
    }
}
