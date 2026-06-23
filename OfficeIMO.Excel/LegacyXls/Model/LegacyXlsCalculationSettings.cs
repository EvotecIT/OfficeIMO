namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Calculation settings parsed from BIFF calculation records.
    /// </summary>
    public sealed class LegacyXlsCalculationSettings {
        private readonly List<LegacyXlsCalculationSettingRecord> _records = new();

        /// <summary>Gets the source calculation records in workbook order.</summary>
        public IReadOnlyList<LegacyXlsCalculationSettingRecord> Records => _records;

        /// <summary>Gets the last parsed calculation mode.</summary>
        public LegacyXlsCalculationMode? Mode { get; private set; }

        /// <summary>Gets the last parsed maximum iteration count.</summary>
        public short? IterationCount { get; private set; }

        /// <summary>Gets whether the workbook uses full calculation precision.</summary>
        public bool? FullPrecision { get; private set; }

        /// <summary>Gets whether formulas use A1 reference style.</summary>
        public bool? A1ReferenceMode { get; private set; }

        /// <summary>Gets the last parsed maximum calculation change for iterative calculation.</summary>
        public double? Delta { get; private set; }

        /// <summary>Gets whether iterative calculation is enabled.</summary>
        public bool? IterationEnabled { get; private set; }

        /// <summary>Gets whether formulas are recalculated before saving.</summary>
        public bool? RecalculateBeforeSave { get; private set; }

        internal void AddRecord(LegacyXlsCalculationSettingRecord record) {
            _records.Add(record);
            switch (record.Kind) {
                case LegacyXlsCalculationSettingKind.IterationCount:
                    IterationCount = record.SignedValue;
                    break;
                case LegacyXlsCalculationSettingKind.Mode:
                    Mode = record.Mode;
                    break;
                case LegacyXlsCalculationSettingKind.FullPrecision:
                    FullPrecision = record.BooleanValue;
                    break;
                case LegacyXlsCalculationSettingKind.A1ReferenceMode:
                    A1ReferenceMode = record.BooleanValue;
                    break;
                case LegacyXlsCalculationSettingKind.Delta:
                    Delta = record.DoubleValue;
                    break;
                case LegacyXlsCalculationSettingKind.IterationEnabled:
                    IterationEnabled = record.BooleanValue;
                    break;
                case LegacyXlsCalculationSettingKind.RecalculateBeforeSave:
                    RecalculateBeforeSave = record.BooleanValue;
                    break;
            }
        }
    }
}
