namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader : IDataReaderFastMappingValues, IDataReaderTypedMappingCompatibility {
        bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

        bool IDataReaderTypedMappingCompatibility.CanUseTypedGetter(int ordinal, Type targetType) {
            ValidateReadableOrdinal(ordinal);
            if (_options.NumericAsDecimal || _options.CellValueConverter != null) return false;

            ValueKind kind = _kinds[ordinal];
            if (targetType == typeof(string)) {
                // Numeric/date formatting and text conversion keep the ordinary mapping rules.
                return kind is ValueKind.Text or ValueKind.Error;
            }
            if (targetType == typeof(bool)) return kind == ValueKind.Boolean;
            if (targetType == typeof(DateTime)) return kind == ValueKind.Date;

            // Date-to-number mapping retains the original serial and converter input.
            return kind == ValueKind.Number
                && (targetType == typeof(byte) || targetType == typeof(short)
                    || targetType == typeof(int) || targetType == typeof(long)
                    || targetType == typeof(float) || targetType == typeof(double)
                    || targetType == typeof(decimal));
        }
    }
}
