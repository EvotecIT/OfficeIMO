using OfficeIMO.Excel.LegacyXls.Projection;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
        bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

        bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) {
            ValidateReadableOrdinal(ordinal);
            if (_options.NumericAsDecimal || _options.CellValueConverter != null) return false;

            XlsbTabularValueKind kind = _kinds[ordinal];
            if (targetType == typeof(string)) {
                // Numeric/date formatting and text conversion keep the ordinary mapping rules.
                return kind is XlsbTabularValueKind.Text or XlsbTabularValueKind.Error;
            }
            if (targetType == typeof(bool)) return kind == XlsbTabularValueKind.Boolean;
            if (targetType == typeof(DateTime)) {
                // Invalid date serials must still fail after the model is constructed.
                return kind is XlsbTabularValueKind.Date or XlsbTabularValueKind.Time
                    && LegacyXlsDateSerialConverter.TryConvert(
                        _numbers[ordinal], _uses1904DateSystem, out _, kind == XlsbTabularValueKind.Date);
            }

            // Date-to-number mapping retains serial validation and the original converter input.
            return kind == XlsbTabularValueKind.Number
                && (targetType == typeof(byte) || targetType == typeof(short)
                    || targetType == typeof(int) || targetType == typeof(long)
                    || targetType == typeof(float) || targetType == typeof(double)
                    || targetType == typeof(decimal));
        }
    }
}
