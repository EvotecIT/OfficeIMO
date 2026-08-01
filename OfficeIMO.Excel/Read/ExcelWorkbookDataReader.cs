#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Excel.Xlsb.Read;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Package-owned ADO.NET projection for XLSX, XLSM, and XLSB workbook worksheets.
    /// </summary>
    public sealed class ExcelWorkbookDataReader : DbDataReader {
        private readonly IReadOnlyList<SheetSelection> _sheets;
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly Func<int, DbDataReader> _openSheet;
        private readonly IDisposable _owner;
        private readonly CultureInfo _culture;
        private DbDataReader _current;
        private int _resultIndex;
        private bool _closed;

        private ExcelWorkbookDataReader(
            IReadOnlyList<SheetSelection> sheets,
            Func<int, DbDataReader> openSheet,
            IDisposable owner,
            CultureInfo culture) {
            if (sheets.Count == 0) {
                owner.Dispose();
                throw new InvalidDataException("The workbook contains no readable worksheets.");
            }

            _sheets = sheets;
            _sheetNames = sheets.Select(static sheet => sheet.Name).ToArray();
            _openSheet = openSheet;
            _owner = owner;
            _culture = culture;
            _current = _openSheet(0);
        }

        internal static ExcelWorkbookDataReader OpenOpenXml(string path, ExcelReadOptions options) =>
            CreateOpenXml(ExcelDocumentReader.Open(path, options), options);

        internal static ExcelWorkbookDataReader OpenOpenXml(byte[] bytes, ExcelReadOptions options) =>
            CreateOpenXml(ExcelDocumentReader.Open(bytes, options), options);

        internal static ExcelWorkbookDataReader WrapOpenXml(ExcelDocumentReader owner, ExcelReadOptions options) =>
            CreateOpenXml(owner, options);

        internal static ExcelWorkbookDataReader OpenBinary(string path, ExcelReadOptions options) =>
            string.IsNullOrWhiteSpace(options.A1Range)
                ? CreateBinary(XlsbTabularWorkbook.Open(path, options, options.CancellationToken), options)
                : OpenProjectedBinary(path, options);

        internal static ExcelWorkbookDataReader OpenBinary(byte[] bytes, ExcelReadOptions options) =>
            string.IsNullOrWhiteSpace(options.A1Range)
                ? CreateBinary(XlsbTabularWorkbook.Open(bytes, options, options.CancellationToken), options)
                : OpenProjectedBinary(bytes, options);

        private static ExcelWorkbookDataReader OpenProjectedBinary(string path, ExcelReadOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            ExcelDocument document = ExcelDocument.Load(path, CreateProjectedBinaryLoadOptions(options));
            try {
                options.CancellationToken.ThrowIfCancellationRequested();
            } catch {
                document.Dispose();
                throw;
            }
            return WrapProjectedBinary(document, options);
        }

        private static ExcelWorkbookDataReader OpenProjectedBinary(byte[] bytes, ExcelReadOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            using var stream = new MemoryStream(bytes, writable: false);
            ExcelDocument document = ExcelDocument.Load(stream, CreateProjectedBinaryLoadOptions(options));
            try {
                options.CancellationToken.ThrowIfCancellationRequested();
            } catch {
                document.Dispose();
                throw;
            }
            return WrapProjectedBinary(document, options);
        }

        private static ExcelLoadOptions CreateProjectedBinaryLoadOptions(ExcelReadOptions options) =>
            new() {
                AccessMode = DocumentAccessMode.ReadOnly,
                PersistenceMode = DocumentPersistenceMode.Explicit,
                MaxInputBytes = options.MaxInputBytes,
                XlsbImportOptions = new OfficeIMO.Excel.Xlsb.XlsbImportOptions {
                    MaxPackageBytes = options.MaxInputBytes,
                    MaxCells = options.MaxXlsbCells,
                    MaxSharedStrings = options.MaxSharedStringItems,
                    ReportPreservedRecords = false
                }
            };

        private static ExcelWorkbookDataReader WrapProjectedBinary(ExcelDocument document, ExcelReadOptions options) {
            ExcelDocumentReader? readerOwner = null;
            try {
                readerOwner = ExcelDocumentReader.Wrap(document._spreadSheetDocument, options);
                return CreateOpenXml(readerOwner, options, document);
            } catch {
                if (readerOwner == null) {
                    document.Dispose();
                }
                throw;
            }
        }

        internal static ExcelWorkbookDataReader OpenLegacy(string path, ExcelReadOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = OfficeIMO.Drawing.Internal.OfficeStreamReader.ReadRemainingBytes(
                stream,
                options.CancellationToken,
                options.MaxInputBytes);
            return OpenLegacyCore(bytes, path, options);
        }

        internal static ExcelWorkbookDataReader OpenLegacy(byte[] bytes, ExcelReadOptions options) {
            return OpenLegacyCore(bytes, sourcePath: null, options);
        }

        private static ExcelWorkbookDataReader OpenLegacyCore(
            byte[] bytes,
            string? sourcePath,
            ExcelReadOptions options) {
            LegacyXlsImportOptions importOptions = CreateLegacyImportOptions(options);
            return CreateLegacy(
                ExcelDocument.LoadLegacyXlsFromNormalFlow(
                    bytes,
                    readOnly: true,
                    saveOnDispose: false,
                    sourcePath,
                    importOptions,
                    workbook => ValidateLegacyFormulaProjection(workbook, options)),
                options);
        }

        internal static void ValidateLegacyFormulaProjection(
            LegacyXlsWorkbook workbook,
            ExcelReadOptions options) {
            if (options.UseCachedFormulaResult) {
                return;
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<SheetSelection> sheets = SelectSheets(
                workbook.Worksheets.Select(static worksheet => worksheet.Name).ToArray(),
                options);
            (int r1, int c1, int r2, int c2)? range = string.IsNullOrWhiteSpace(options.A1Range)
                ? null
                : A1.ParseRange(options.A1Range!);
            foreach (SheetSelection selection in sheets) {
                options.CancellationToken.ThrowIfCancellationRequested();
                LegacyXlsWorksheet worksheet = workbook.Worksheets[selection.WorkbookIndex];

                foreach (LegacyXlsCell cell in worksheet.Cells) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    if (range.HasValue
                        && (cell.Row < range.Value.r1
                            || cell.Row > range.Value.r2
                            || cell.Column < range.Value.c1
                            || cell.Column > range.Value.c2)) {
                        continue;
                    }
                    if (!cell.IsFormula
                        || (!string.IsNullOrWhiteSpace(cell.FormulaText)
                            && LegacyXlsWorkbookProjector.ShouldProjectFormula(workbook, cell.FormulaText!))) {
                        continue;
                    }

                    string reference = A1.CellReference(cell.Row, cell.Column);
                    throw new NotSupportedException(
                        $"Legacy XLS formula text cannot be projected for '{worksheet.Name}'!{reference}. " +
                        "UseCachedFormulaResult=false cannot return the cached value as ordinary data.");
                }
            }
        }

        internal static LegacyXlsImportOptions CreateLegacyImportOptions(ExcelReadOptions options) =>
            new() {
                MaxInputBytes = options.MaxInputBytes > int.MaxValue
                    ? int.MaxValue
                    : checked((int)options.MaxInputBytes),
                CancellationToken = options.CancellationToken
            };

        private static ExcelWorkbookDataReader CreateLegacy(
            ExcelDocument document,
            ExcelReadOptions options) {
            ExcelDocumentReader? readerOwner = null;
            try {
                options.CancellationToken.ThrowIfCancellationRequested();
                readerOwner = ExcelDocumentReader.Wrap(document._spreadSheetDocument, options);
                return CreateOpenXml(readerOwner, options, document);
            } catch {
                if (readerOwner == null) {
                    document.Dispose();
                }
                throw;
            }
        }

        private static ExcelWorkbookDataReader CreateOpenXml(
            ExcelDocumentReader owner,
            ExcelReadOptions options,
            IDisposable? additionalOwner = null) {
            IDisposable lifetime = additionalOwner == null
                ? owner
                : new CompositeOwner(owner, additionalOwner);
            try {
                IReadOnlyList<string> availableSheets = owner.GetValidatedWorksheetNames();
                ValidateUniqueSheetNames(availableSheets, options.CancellationToken);
                IReadOnlyList<SheetSelection> sheets = SelectSheets(availableSheets, options);
                return new ExcelWorkbookDataReader(
                    sheets,
                    index => OpenOpenXmlSheet(owner, sheets[index].Name, options),
                    lifetime,
                    options.Culture);
            } catch {
                lifetime.Dispose();
                throw;
            }
        }

        private static DbDataReader OpenOpenXmlSheet(
            ExcelDocumentReader owner,
            string sheetName,
            ExcelReadOptions options) {
            ExcelSheetReader sheet = owner.GetSheet(sheetName);
            return string.IsNullOrWhiteSpace(options.A1Range)
                ? (DbDataReader)sheet.ReadUsedRangeAsDataReader(
                    headersInFirstRow: options.HasHeaderRow,
                    schemaSampleRows: options.InferSchema ? options.SchemaSampleRows : 0,
                    ct: options.CancellationToken)
                : (DbDataReader)sheet.ReadRangeAsDataReader(
                    options.A1Range!,
                    headersInFirstRow: options.HasHeaderRow,
                    chunkRows: Math.Min(1024, options.MaxDataReaderChunkRows),
                    schemaSampleRows: options.InferSchema ? options.SchemaSampleRows : 0,
                    ct: options.CancellationToken);
        }

        private static ExcelWorkbookDataReader CreateBinary(
            XlsbTabularWorkbook owner,
            ExcelReadOptions options) {
            try {
                IReadOnlyList<SheetSelection> sheets = SelectSheets(owner.TableNames, options);
                return new ExcelWorkbookDataReader(
                    sheets,
                    index => owner.OpenTable(
                        sheets[index].Name,
                        options.HasHeaderRow,
                        options,
                        options.CancellationToken),
                    owner,
                    options.Culture);
            } catch {
                owner.Dispose();
                throw;
            }
        }

        private static void ValidateUniqueSheetNames(
            IReadOnlyList<string> sheetNames,
            CancellationToken cancellationToken) {
            var uniqueNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string sheetName in sheetNames) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!uniqueNames.Add(sheetName)) {
                    throw new InvalidDataException(
                        $"The workbook contains duplicate worksheet name '{sheetName}' under case-insensitive matching.");
                }
            }
        }

        private static IReadOnlyList<SheetSelection> SelectSheets(
            IReadOnlyList<string> sheetNames,
            ExcelReadOptions options) {
            if (!string.IsNullOrWhiteSpace(options.SheetName) && options.SheetIndex.HasValue) {
                throw new ArgumentException("SheetName and SheetIndex cannot be used together.", nameof(options));
            }
            if (options.SheetIndex.HasValue) {
                int selectedIndex = options.SheetIndex.Value;
                if (selectedIndex < 0 || selectedIndex >= sheetNames.Count) {
                    throw new ArgumentOutOfRangeException(nameof(options), $"Sheet index {selectedIndex} is outside the workbook worksheet range.");
                }
                return new[] { new SheetSelection(sheetNames[selectedIndex], selectedIndex) };
            }
            if (string.IsNullOrWhiteSpace(options.SheetName)) {
                return sheetNames
                    .Select(static (name, index) => new SheetSelection(name, index))
                    .ToArray();
            }

            int matchIndex = -1;
            for (int index = 0; index < sheetNames.Count; index++) {
                if (string.Equals(sheetNames[index], options.SheetName, StringComparison.OrdinalIgnoreCase)) {
                    matchIndex = index;
                    break;
                }
            }
            if (matchIndex < 0) {
                throw new KeyNotFoundException($"Sheet '{options.SheetName}' was not found.");
            }

            return new[] { new SheetSelection(sheetNames[matchIndex], matchIndex) };
        }

        private readonly record struct SheetSelection(string Name, int WorkbookIndex);

        /// <summary>Gets worksheet names exposed by this reader in workbook order.</summary>
        public IReadOnlyList<string> SheetNames => _sheetNames;

        /// <summary>Gets the zero-based workbook index of the current worksheet.</summary>
        public int CurrentSheetIndex => _sheets[_resultIndex].WorkbookIndex;

        /// <summary>Gets the zero-based index of the current reader result within <see cref="SheetNames"/>.</summary>
        public int CurrentResultIndex => _resultIndex;

        /// <summary>Gets the worksheet name for the current result.</summary>
        public string CurrentSheetName => _sheetNames[_resultIndex];

        /// <inheritdoc />

        public override bool NextResult() {
            ThrowIfClosed();
            if (_resultIndex + 1 >= _sheetNames.Count) {
                return false;
            }

            _current.Dispose();
            int nextResultIndex = _resultIndex + 1;
            try {
                _current = _openSheet(nextResultIndex);
            } catch {
                CloseAfterSheetOpenFailure();
                throw;
            }
            _resultIndex = nextResultIndex;
            return true;
        }

        private void CloseAfterSheetOpenFailure() {
            _closed = true;
            try {
                _current.Dispose();
            } catch {
                // Preserve the worksheet-open failure while attempting complete cleanup.
            }
            try {
                _owner.Dispose();
            } catch {
                // Preserve the worksheet-open failure while attempting complete cleanup.
            }
        }

        /// <inheritdoc />

        public override void Close() {
            if (_closed) {
                return;
            }

            _closed = true;
            try {
                _current.Dispose();
            } finally {
                _owner.Dispose();
            }
        }

        /// <inheritdoc />

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Close();
            }
            base.Dispose(disposing);
        }

        /// <inheritdoc />

        public override object this[int ordinal] => _current[ordinal];
        /// <inheritdoc />
        public override object this[string name] => _current[name];
        /// <inheritdoc />
        public override int Depth => _current.Depth;
        /// <inheritdoc />
        public override int FieldCount => _current.FieldCount;
        /// <inheritdoc />
        public override bool HasRows => _current.HasRows;
        /// <inheritdoc />
        public override bool IsClosed => _closed || _current.IsClosed;
        /// <inheritdoc />
        public override int RecordsAffected => _current.RecordsAffected;
        /// <inheritdoc />
        public override int VisibleFieldCount => _current.VisibleFieldCount;
        /// <inheritdoc />
        public override bool GetBoolean(int ordinal) => _current.GetBoolean(ordinal);
        /// <inheritdoc />
        public override byte GetByte(int ordinal) => _current.GetByte(ordinal);
        /// <inheritdoc />
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
            _current.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
        /// <inheritdoc />
        public override char GetChar(int ordinal) => _current.GetChar(ordinal);
        /// <inheritdoc />
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
            _current.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
        /// <inheritdoc />
        public override string GetDataTypeName(int ordinal) => _current.GetDataTypeName(ordinal);
        /// <inheritdoc />
        public override DateTime GetDateTime(int ordinal) => _current.GetDateTime(ordinal);
        /// <inheritdoc />
        public override decimal GetDecimal(int ordinal) => _current.GetDecimal(ordinal);
        /// <inheritdoc />
        public override double GetDouble(int ordinal) => _current.GetDouble(ordinal);
        /// <inheritdoc />
        public override IEnumerator GetEnumerator() => _current.GetEnumerator();
        /// <inheritdoc />
#if NET8_0_OR_GREATER
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
        public override Type GetFieldType(int ordinal) => _current.GetFieldType(ordinal);

        /// <inheritdoc />

        public override T GetFieldValue<T>(int ordinal) {
            Type destinationType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            if (destinationType == typeof(string)) return (T)(object)GetString(ordinal);
            if (destinationType == typeof(bool)) return (T)(object)GetBoolean(ordinal);
            if (destinationType == typeof(byte)) return (T)(object)GetByte(ordinal);
            if (destinationType == typeof(short)) return (T)(object)GetInt16(ordinal);
            if (destinationType == typeof(int)) return (T)(object)GetInt32(ordinal);
            if (destinationType == typeof(long)) return (T)(object)GetInt64(ordinal);
            if (destinationType == typeof(float)) return (T)(object)GetFloat(ordinal);
            if (destinationType == typeof(double)) return (T)(object)GetDouble(ordinal);
            if (destinationType == typeof(decimal)) return (T)(object)GetDecimal(ordinal);
            if (destinationType == typeof(DateTime)) return (T)(object)GetDateTime(ordinal);
            if (destinationType == typeof(Guid)) return (T)(object)GetGuid(ordinal);

            object value = GetValue(ordinal);
            if (value is T typed) {
                return typed;
            }

            return (T)Convert.ChangeType(value, destinationType, _culture);
        }

        /// <inheritdoc />

        public override float GetFloat(int ordinal) => _current.GetFloat(ordinal);
        /// <inheritdoc />
        public override Guid GetGuid(int ordinal) => _current.GetGuid(ordinal);
        /// <inheritdoc />
        public override short GetInt16(int ordinal) => _current.GetInt16(ordinal);
        /// <inheritdoc />
        public override int GetInt32(int ordinal) => _current.GetInt32(ordinal);
        /// <inheritdoc />
        public override long GetInt64(int ordinal) => _current.GetInt64(ordinal);
        /// <inheritdoc />
        public override string GetName(int ordinal) => _current.GetName(ordinal);
        /// <inheritdoc />
        public override int GetOrdinal(string name) => _current.GetOrdinal(name);
        /// <inheritdoc />
        public override string GetString(int ordinal) => _current.GetString(ordinal);
        /// <inheritdoc />
        public override object GetValue(int ordinal) => _current.GetValue(ordinal);
        /// <inheritdoc />
        public override int GetValues(object[] values) => _current.GetValues(values);
        /// <inheritdoc />
        public override bool IsDBNull(int ordinal) => _current.IsDBNull(ordinal);
        /// <inheritdoc />
        public override bool Read() {
            ThrowIfClosed();
            return _current.Read();
        }
        /// <inheritdoc />
        public override DataTable? GetSchemaTable() => _current.GetSchemaTable();

        private void ThrowIfClosed() {
            if (IsClosed) {
                throw new InvalidOperationException("The Excel data reader is closed.");
            }
        }

        private sealed class CompositeOwner : IDisposable {
            private IDisposable? _primary;
            private IDisposable? _secondary;

            internal CompositeOwner(IDisposable primary, IDisposable secondary) {
                _primary = primary;
                _secondary = secondary;
            }

            public void Dispose() {
                IDisposable? primary = Interlocked.Exchange(ref _primary, null);
                IDisposable? secondary = Interlocked.Exchange(ref _secondary, null);
                try {
                    primary?.Dispose();
                } finally {
                    secondary?.Dispose();
                }
            }
        }
    }
}
