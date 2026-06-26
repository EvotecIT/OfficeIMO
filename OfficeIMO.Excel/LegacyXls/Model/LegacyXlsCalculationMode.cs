namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Calculation mode recorded in a legacy BIFF workbook.
    /// </summary>
    public enum LegacyXlsCalculationMode {
        /// <summary>Formulas recalculate only when requested.</summary>
        Manual = 0,

        /// <summary>Formulas recalculate automatically.</summary>
        Automatic = 1,

        /// <summary>Formulas recalculate automatically except for data tables.</summary>
        AutomaticExceptTables = 2
    }
}
