using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reading options controlling conversion behavior and execution policy.
    /// </summary>
    public sealed class ExcelReadOptions {
        private int _maxSharedStringItems = 1_000_000;
        private int _maxSharedStringItemCharacters = 32_767;
        private long _maxSharedStringCharacters = 64L * 1024L * 1024L;
        private long _maxInputBytes = 512L * 1024L * 1024L;

        /// <summary>Maximum workbook bytes buffered by <see cref="ExcelDocumentReader"/>. Default: 512 MiB.</summary>
        public long MaxInputBytes {
            get => _maxInputBytes;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Workbook input limit must be greater than zero.");
                }

                _maxInputBytes = value;
            }
        }

        /// <summary>Maximum columns exposed by one range data reader.</summary>
        public int MaxDataReaderColumns { get; set; } = 16_384;

        /// <summary>Maximum worksheet rows materialized in one data-reader chunk.</summary>
        public int MaxDataReaderChunkRows { get; set; } = 8_192;

        /// <summary>Maximum rows retained for data-reader schema inference.</summary>
        public int MaxDataReaderSchemaSampleRows { get; set; } = 4_096;

        /// <summary>Maximum cells materialized by a data-reader chunk or schema sample.</summary>
        public long MaxDataReaderBufferedCells { get; set; } = 1_000_000L;

        /// <summary>Maximum cells materialized by one dense range read. Default: 1,000,000.</summary>
        public long MaxRangeCells { get; set; } = 1_000_000L;

        /// <summary>Maximum out-of-order rows retained by one typed streaming read. Default: 8,192.</summary>
        public int MaxPendingTypedRows { get; set; } = 8_192;

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
        /// When true, numeric cells are returned as decimal where possible; otherwise double is used.
        /// </summary>
        public bool NumericAsDecimal { get; set; } = false;

        /// <summary>
        /// When true, DataTable reads infer stable column types from the materialized range.
        /// Mixed-type columns stay object-typed.
        /// </summary>
        public bool InferDataTableColumnTypes { get; set; } = true;

        /// <summary>
        /// When true, typed object readers throw if selected headers cannot be mapped
        /// deterministically to writable properties.
        /// </summary>
        public bool StrictTypedMapping { get; set; } = false;

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
        /// Maximum number of entries loaded from the workbook shared-string table.
        /// This protects readers from malformed workbooks that advertise or contain
        /// unbounded shared-string tables.
        /// </summary>
        public int MaxSharedStringItems {
            get => _maxSharedStringItems;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Shared-string item limit must be greater than zero.");
                }

                _maxSharedStringItems = value;
            }
        }

        /// <summary>
        /// Maximum character length for one shared-string item. The default matches
        /// Excel's worksheet cell text limit.
        /// </summary>
        public int MaxSharedStringItemCharacters {
            get => _maxSharedStringItemCharacters;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Shared-string item character limit must be greater than zero.");
                }

                _maxSharedStringItemCharacters = value;
            }
        }

        /// <summary>
        /// Maximum aggregate characters loaded from the shared-string table.
        /// </summary>
        public long MaxSharedStringCharacters {
            get => _maxSharedStringCharacters;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Shared-string aggregate character limit must be greater than zero.");
                }

                _maxSharedStringCharacters = value;
            }
        }

        /// <summary>
        /// Initializes reading defaults and per-operation thresholds.
        /// </summary>
        public ExcelReadOptions() {
            Execution.OperationThresholds["ReadRange"] = 100_000;
            Execution.OperationThresholds["ReadRangeAsDataTable"] = 100_000;
            Execution.OperationThresholds["ReadObjects"] = 10_000;
            Execution.OperationThresholds["ReadObjectsAs"] = 100_000;
            Execution.OperationThresholds["ReadRangeStream"] = 100_000;
            Execution.OperationThresholds["ReadRows"] = 20_000;
        }
    }
}
