namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded SerAuxErrBar chart error-bar options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartErrorBarOptions {
        internal LegacyXlsChartErrorBarOptions(byte direction, byte valueSource, bool hasTeeTop, byte reserved, double value, ushort customValueCount) {
            Direction = direction;
            DirectionName = GetDirectionName(direction);
            ValueSource = valueSource;
            ValueSourceName = GetValueSourceName(valueSource);
            HasTeeTop = hasTeeTop;
            Reserved = reserved;
            Value = value;
            CustomValueCount = customValueCount;
        }

        /// <summary>Gets the raw error-bar direction value.</summary>
        public byte Direction { get; }

        /// <summary>Gets the decoded error-bar direction name.</summary>
        public string DirectionName { get; }

        /// <summary>Gets whether the error-bar direction is defined by MS-XLS.</summary>
        public bool HasKnownDirection => Direction >= 0x01 && Direction <= 0x04;

        /// <summary>Gets the raw error-bar value-source value.</summary>
        public byte ValueSource { get; }

        /// <summary>Gets the decoded error-bar value-source name.</summary>
        public string ValueSourceName { get; }

        /// <summary>Gets whether the error-bar value source is defined by MS-XLS.</summary>
        public bool HasKnownValueSource => ValueSource >= 0x01 && ValueSource <= 0x05;

        /// <summary>Gets whether the error bars are T-shaped.</summary>
        public bool HasTeeTop { get; }

        /// <summary>Gets the reserved byte, which is expected to be 0x01.</summary>
        public byte Reserved { get; }

        /// <summary>Gets whether the reserved byte has the expected value.</summary>
        public bool HasExpectedReservedValue => Reserved == 0x01;

        /// <summary>Gets the fixed value, percentage, or number of standard deviations for non-custom error bars.</summary>
        public double Value { get; }

        /// <summary>Gets the number of values or references used for custom error bars.</summary>
        public ushort CustomValueCount { get; }

        /// <summary>Gets whether the value field is meaningful for the decoded value source.</summary>
        public bool UsesValue => ValueSource != 0x04 && ValueSource != 0x05;

        /// <summary>Gets whether the custom value count field is meaningful for the decoded value source.</summary>
        public bool UsesCustomValueCount => ValueSource == 0x04;

        private static string GetDirectionName(byte value) {
            switch (value) {
                case 0x01:
                    return "HorizontalPlus";
                case 0x02:
                    return "HorizontalMinus";
                case 0x03:
                    return "VerticalPlus";
                case 0x04:
                    return "VerticalMinus";
                default:
                    return $"Direction:0x{value:X2}";
            }
        }

        private static string GetValueSourceName(byte value) {
            switch (value) {
                case 0x01:
                    return "Percentage";
                case 0x02:
                    return "FixedValue";
                case 0x03:
                    return "StandardDeviation";
                case 0x04:
                    return "CustomValues";
                case 0x05:
                    return "StandardError";
                default:
                    return $"ValueSource:0x{value:X2}";
            }
        }
    }
}
