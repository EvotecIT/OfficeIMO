using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Reading options controlling conversion behavior and execution policy.
    /// </summary>
    public sealed class ExcelReadOptions {
        private int _maxSharedStringItems = 1_000_000;
        private int _maxSharedStringItemCharacters = 32_767;
        private long _maxSharedStringCharacters = 64L * 1024L * 1024L;
        private long _maxInputBytes = 512L * 1024L * 1024L;
        private int _schemaSampleRows = 1_024;
        private int _maxXlsbCells = 4_000_000;

        /// <summary>
        /// Gets or sets the worksheet exposed by <see cref="ExcelDocument.OpenDataReader(string, ExcelReadOptions?)"/>.
        /// When omitted, worksheets are exposed in workbook order through
        /// <see cref="System.Data.Common.DbDataReader.NextResult"/>.
        /// </summary>
        public string? SheetName { get; set; }

        /// <summary>Gets or sets whether the first row supplies column names.</summary>
        public bool HasHeaderRow { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the data-reader schema is inferred from worksheet values.
        /// Native cell values are preserved regardless of this setting.
        /// </summary>
        public bool InferSchema { get; set; }

        /// <summary>Gets or sets the maximum rows sampled when schema inference is enabled.</summary>
        public int SchemaSampleRows {
            get => _schemaSampleRows;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Schema sample rows must be greater than zero.");
                }

                _schemaSampleRows = value;
            }
        }

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

        /// <summary>
        /// Maximum populated cell records accepted across one XLSB workbook. Default: 4,000,000.
        /// This is an aggregate safety limit, independent of per-chunk data-reader buffering.
        /// </summary>
        public int MaxXlsbCells {
            get => _maxXlsbCells;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "XLSB cell limit must be greater than zero.");
                }

                _maxXlsbCells = value;
            }
        }

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
        /// Cancellation observed while opening and streaming workbook data.
        /// </summary>
        public CancellationToken CancellationToken { get; set; }

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
