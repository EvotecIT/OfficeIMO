namespace OfficeIMO.Odbc {
    /// <summary>Defines the supported input binding pairs without emulating general ODBC conversions.</summary>
    internal static class NativeParameter {
        internal static short DefaultCType(short sqlType) => sqlType switch {
            1 or 12 or -1 => 1, -8 or -9 or -10 => -8,
            -7 => -7, 5 => 5, 4 => 4, -5 => -25, 7 => 7, 8 => 8,
            91 => 91, 93 => 93, _ => 0
        };

        internal static bool Supports(short type, short sqlType) => sqlType switch {
            1 or 12 or -1 or -8 or -9 or -10 => type is 1 or -8,
            5 => type is 5 or -15,
            4 => type is 4 or -16,
            _ => type == DefaultCType(sqlType)
        };

        internal static object? CheckValue(ParameterBinding binding, object? value) {
            if (value is double number && !double.IsFinite(number) || value is float single && !float.IsFinite(single))
                throw new ParameterValueException("22003", "Floating-point parameters must be finite.");
            if (value is string text && binding.Size != 0 && (nuint)text.Length > binding.Size)
                throw new ParameterValueException("22001", "Parameter text exceeds its declared SQL column size.");
            if (value is DateTime timestamp && binding.SqlType == 93) {
                long unit = 1;
                for (int i = binding.Scale; i < 7; i++) unit *= 10;
                if (timestamp.Ticks % unit != 0)
                    throw new ParameterValueException("22008", "Timestamp cannot be represented at the declared fractional precision.");
            }
            return value;
        }
    }

    internal sealed class ParameterValueException(string state, string message) : Exception(message) {
        internal string State { get; } = state;
    }
}
