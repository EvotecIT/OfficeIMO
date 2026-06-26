namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only Units chart metadata.
    /// </summary>
    public sealed class LegacyXlsChartUnits {
        internal LegacyXlsChartUnits(ushort reserved) {
            Reserved = reserved;
        }

        /// <summary>Gets the reserved Units value, which should be zero and ignored.</summary>
        public ushort Reserved { get; }

        /// <summary>Gets whether the reserved Units value is zero.</summary>
        public bool HasZeroReservedValue => Reserved == 0;
    }
}
