namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies a calculation setting parsed from a BIFF calculation record.
    /// </summary>
    public enum LegacyXlsCalculationSettingKind {
        /// <summary>Maximum number of iterations for iterative calculation.</summary>
        IterationCount,

        /// <summary>Workbook calculation mode.</summary>
        Mode,

        /// <summary>Whether formulas are calculated with full precision.</summary>
        FullPrecision,

        /// <summary>Whether formulas use A1 reference style.</summary>
        A1ReferenceMode,

        /// <summary>Maximum calculation change for iterative calculation.</summary>
        Delta,

        /// <summary>Whether iterative calculation is enabled.</summary>
        IterationEnabled,

        /// <summary>Whether formulas are recalculated before saving.</summary>
        RecalculateBeforeSave
    }
}
