namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents workbook locale identifiers decoded from a legacy Country record.
    /// </summary>
    public sealed class LegacyXlsCountryInfo {
        /// <summary>
        /// Creates country metadata.
        /// </summary>
        /// <param name="defaultCountryCode">Country or region code for built-in workbook settings.</param>
        /// <param name="systemCountryCode">Country or region code from the system that saved the workbook.</param>
        public LegacyXlsCountryInfo(ushort defaultCountryCode, ushort systemCountryCode) {
            DefaultCountryCode = defaultCountryCode;
            SystemCountryCode = systemCountryCode;
        }

        /// <summary>Gets the country or region code for built-in workbook settings.</summary>
        public ushort DefaultCountryCode { get; }

        /// <summary>Gets the country or region code from the system that saved the workbook.</summary>
        public ushort SystemCountryCode { get; }
    }
}
