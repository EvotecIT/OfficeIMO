using System.Globalization;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Reading options controlling conversion behavior and execution policy.
    /// </summary>
    public sealed class ExcelReadOptions
    {
        /// <summary>
        /// Execution policy used to decide Sequential vs Parallel conversion.
        /// Reuses the writer-side policy for symmetry.
        /// </summary>
        public OfficeIMO.Excel.ExecutionPolicy Execution { get; } = new();

        /// <summary>
        /// Use cached formula results when present; otherwise returns the formula string.
        /// </summary>
        public bool UseCachedFormulaResult { get; set; } = true;

        /// <summary>
        /// Interpret numeric cells with a date-like number format as DateTime (OADate).
        /// </summary>
        public bool TreatDatesUsingNumberFormat { get; set; } = true;

        /// <summary>
        /// Culture used when parsing numbers and dates stored as strings.
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

        /// <summary>
        /// When true, matrix/range readers fill unspecified cells with nulls.
        /// </summary>
        public bool FillBlanksInRanges { get; set; } = true;

        /// <summary>
        /// Normalize headers for object mapping by trimming and collapsing whitespace.
        /// </summary>
        public bool NormalizeHeaders { get; set; } = true;

        /// <summary>
        /// Optional cell-level converter hook. If provided and it returns a handled value,
        /// the built-in conversion pipeline is skipped and the returned value is used.
        /// </summary>
        public Func<ExcelCellContext, ExcelCellValue>? CellValueConverter { get; set; }

        /// <summary>
        /// Optional type conversion hook used by typed readers (ReadColumnAs/ReadRowsAs/ReadRangeAs and object mapping).
        /// If it returns ok=true, its value is used; otherwise the built-in converter is used.
        /// </summary>
        public Func<object, Type, CultureInfo, (bool ok, object? value)>? TypeConverter { get; set; }

        /// <summary>
        /// Initializes reading defaults and per-operation thresholds.
        /// </summary>
        public ExcelReadOptions()
        {
            Execution.OperationThresholds["ReadRange"] = 10_000;
            Execution.OperationThresholds["ReadObjects"] = 2_000;
            Execution.OperationThresholds["ReadRows"] = 20_000;
        }
    }
}
