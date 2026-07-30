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
    internal sealed class ExcelWorkbookDataReader : DbDataReader {
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly Func<int, DbDataReader> _openSheet;
        private readonly IDisposable _owner;
        private readonly CultureInfo _culture;
        private DbDataReader _current;
        private int _sheetIndex;
        private bool _closed;

        private ExcelWorkbookDataReader(
            IReadOnlyList<string> sheetNames,
            Func<int, DbDataReader> openSheet,
            IDisposable owner,
            CultureInfo culture) {
            if (sheetNames.Count == 0) {
                owner.Dispose();
                throw new InvalidDataException("The workbook contains no readable worksheets.");
            }

            _sheetNames = sheetNames;
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
            CreateBinary(XlsbTabularWorkbook.Open(path, options, options.CancellationToken), options);

        internal static ExcelWorkbookDataReader OpenBinary(byte[] bytes, ExcelReadOptions options) =>
            CreateBinary(XlsbTabularWorkbook.Open(bytes, options, options.CancellationToken), options);

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
            foreach (LegacyXlsWorksheet worksheet in workbook.Worksheets) {
                foreach (LegacyXlsCell cell in worksheet.Cells) {
                    options.CancellationToken.ThrowIfCancellationRequested();
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
                IReadOnlyList<string> sheets = SelectSheets(owner.GetSheetNames(), options.SheetName);
                return new ExcelWorkbookDataReader(
                    sheets,
                    index => OpenOpenXmlSheet(owner, sheets[index], options),
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
            sheet.ValidateFormulaTextDataReaderProjection(options.CancellationToken);
            return (DbDataReader)sheet.ReadUsedRangeAsDataReader(
                headersInFirstRow: options.HasHeaderRow,
                schemaSampleRows: options.InferSchema ? options.SchemaSampleRows : 0,
                ct: options.CancellationToken);
        }

        private static ExcelWorkbookDataReader CreateBinary(
            XlsbTabularWorkbook owner,
            ExcelReadOptions options) {
            try {
                IReadOnlyList<string> sheets = SelectSheets(owner.TableNames, options.SheetName);
                return new ExcelWorkbookDataReader(
                    sheets,
                    index => owner.OpenTable(
                        sheets[index],
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

        private static IReadOnlyList<string> SelectSheets(
            IReadOnlyList<string> sheetNames,
            string? selectedSheet) {
            if (string.IsNullOrWhiteSpace(selectedSheet)) {
                return sheetNames;
            }

            string? match = sheetNames.FirstOrDefault(
                name => string.Equals(name, selectedSheet, StringComparison.OrdinalIgnoreCase));
            if (match == null) {
                throw new KeyNotFoundException($"Sheet '{selectedSheet}' was not found.");
            }

            return new[] { match };
        }

        public override bool NextResult() {
            ThrowIfClosed();
            if (_sheetIndex + 1 >= _sheetNames.Count) {
                return false;
            }

            _current.Dispose();
            int nextSheetIndex = _sheetIndex + 1;
            try {
                _current = _openSheet(nextSheetIndex);
            } catch {
                CloseAfterSheetOpenFailure();
                throw;
            }
            _sheetIndex = nextSheetIndex;
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

        public override void Close() {
            if (_closed) {
                return;
            }

            _closed = true;
            _current.Dispose();
            _owner.Dispose();
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                Close();
            }
            base.Dispose(disposing);
        }

        public override object this[int ordinal] => _current[ordinal];
        public override object this[string name] => _current[name];
        public override int Depth => _current.Depth;
        public override int FieldCount => _current.FieldCount;
        public override bool HasRows => _current.HasRows;
        public override bool IsClosed => _closed || _current.IsClosed;
        public override int RecordsAffected => _current.RecordsAffected;
        public override int VisibleFieldCount => _current.VisibleFieldCount;
        public override bool GetBoolean(int ordinal) => _current.GetBoolean(ordinal);
        public override byte GetByte(int ordinal) => _current.GetByte(ordinal);
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
            _current.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
        public override char GetChar(int ordinal) => _current.GetChar(ordinal);
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
            _current.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
        public override string GetDataTypeName(int ordinal) => _current.GetDataTypeName(ordinal);
        public override DateTime GetDateTime(int ordinal) => _current.GetDateTime(ordinal);
        public override decimal GetDecimal(int ordinal) => _current.GetDecimal(ordinal);
        public override double GetDouble(int ordinal) => _current.GetDouble(ordinal);
        public override IEnumerator GetEnumerator() => _current.GetEnumerator();
#if NET8_0_OR_GREATER
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
        public override Type GetFieldType(int ordinal) => _current.GetFieldType(ordinal);

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

        public override float GetFloat(int ordinal) => _current.GetFloat(ordinal);
        public override Guid GetGuid(int ordinal) => _current.GetGuid(ordinal);
        public override short GetInt16(int ordinal) => _current.GetInt16(ordinal);
        public override int GetInt32(int ordinal) => _current.GetInt32(ordinal);
        public override long GetInt64(int ordinal) => _current.GetInt64(ordinal);
        public override string GetName(int ordinal) => _current.GetName(ordinal);
        public override int GetOrdinal(string name) => _current.GetOrdinal(name);
        public override string GetString(int ordinal) => _current.GetString(ordinal);
        public override object GetValue(int ordinal) => _current.GetValue(ordinal);
        public override int GetValues(object[] values) => _current.GetValues(values);
        public override bool IsDBNull(int ordinal) => _current.IsDBNull(ordinal);
        public override bool Read() {
            ThrowIfClosed();
            return _current.Read();
        }
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
