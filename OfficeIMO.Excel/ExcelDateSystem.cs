namespace OfficeIMO.Excel {
    /// <summary>
    /// Excel workbook date system used for serial date values.
    /// </summary>
    public enum ExcelDateSystem {
        /// <summary>
        /// The default Excel 1900 date system.
        /// </summary>
        NineteenHundred = 1900,

        /// <summary>
        /// The Excel 1904 date system, commonly used by older Mac-originated workbooks.
        /// </summary>
        NineteenFour = 1904,
    }
}
