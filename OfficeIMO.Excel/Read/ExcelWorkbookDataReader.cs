#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Projection;
using OfficeIMO.Excel.LegacyXls.Read;
using OfficeIMO.Excel.Xlsb.Read;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Package-owned ADO.NET projection for XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, and BIFF8 XLS workbook worksheets.
    /// </summary>
    public sealed partial class ExcelWorkbookDataReader : DbDataReader, IDataReaderMappingMetadata, IDataReaderMappingErrorMetadata, IDataReaderFastValueSource {
        private readonly IReadOnlyList<SheetSelection> _sheets;
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly Func<int, CancellationToken, DbDataReader> _openSheet;
        private IDisposable _owner;
        private readonly CultureInfo _culture;

        CultureInfo IDataReaderMappingMetadata.MappingCulture => _culture;

        IReadOnlyList<string>? IDataReaderMappingMetadata.MappingDateTimeFormats => null;
        Func<object, Type, CultureInfo, (bool ok, object? value)>? IDataReaderMappingMetadata.MappingTypeConverter => _typeConverter;
        bool IDataReaderMappingMetadata.RequireAllColumnsMapped => _requireAllColumnsMapped;
        DataMappingErrorValuePolicy IDataReaderMappingErrorMetadata.MappingErrorValuePolicy => _mappingErrorValuePolicy;
        IDataReaderFastValueSource? IDataReaderFastValueSource.FastValueSource =>
            (CurrentReader as IDataReaderFastValueSource)?.FastValueSource;
        bool IDataReaderFastValueSource.TryGetUtf8Value(int ordinal, out ArraySegment<byte> value) {
            if (CurrentReader is IDataReaderFastValueSource source) {
                return source.TryGetUtf8Value(ordinal, out value);
            }

            value = default;
            return false;
        }

        bool IDataReaderFastValueSource.TryGetInt64(int ordinal, out long value) {
            if (CurrentReader is IDataReaderFastValueSource source) {
                return source.TryGetInt64(ordinal, out value);
            }

            value = default;
            return false;
        }

        bool IDataReaderFastValueSource.TryGetDouble(int ordinal, out double value) {
            if (CurrentReader is IDataReaderFastValueSource source) {
                return source.TryGetDouble(ordinal, out value);
            }

            value = default;
            return false;
        }

        bool IDataReaderFastValueSource.TryGetDateTime(int ordinal, out DateTime value) {
            if (CurrentReader is IDataReaderFastValueSource source) {
                return source.TryGetDateTime(ordinal, out value);
            }

            value = default;
            return false;
        }

        private readonly Func<object, Type, CultureInfo, (bool ok, object? value)>? _typeConverter;
        private readonly bool _requireAllColumnsMapped;
        private readonly DataMappingErrorValuePolicy _mappingErrorValuePolicy;
        private readonly CancellationToken _cancellationToken;
        private CancellationTokenSource? _currentSheetCancellationSource;
        private int _resultIndex;
        private bool _closed;

        private ExcelWorkbookDataReader(
            IReadOnlyList<SheetSelection> sheets,
            Func<int, CancellationToken, DbDataReader> openSheet,
            IDisposable owner,
            CultureInfo culture,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            bool requireAllColumnsMapped,
            DataMappingErrorValuePolicy mappingErrorValuePolicy,
            CancellationToken cancellationToken) {
            if (sheets.Count == 0) {
                owner.Dispose();
                throw new InvalidDataException("The workbook contains no readable worksheets.");
            }

            _sheets = sheets;
            _sheetNames = sheets.Select(static sheet => sheet.Name).ToArray();
            _openSheet = openSheet;
            _owner = owner;
            _culture = culture;
            _typeConverter = typeConverter;
            _requireAllColumnsMapped = requireAllColumnsMapped;
            _mappingErrorValuePolicy = mappingErrorValuePolicy;
            _cancellationToken = cancellationToken;

        }

        internal ExcelWorkbookDataReader OwnLifetime(IDisposable lifetime) {
            _owner = new CompositeOwner(_owner, lifetime);
            return this;
        }

        internal static ExcelWorkbookDataReader OpenOpenXml(string path, ExcelReaderOptions options) {
            if (CanUseNativeOpenXmlPath(options)
                && string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                try {
                    return CreateNativeOpenXml(XlsxTabularWorkbook.Open(path, options), options, () => ExcelDocumentReader.Open(path, options));
                } catch (XlsxTabularFastPathNotSupportedException) {
                    // Unusual package and worksheet shapes retain the complete SDK path.
                }
            }

            return CreateOpenXml(ExcelDocumentReader.Open(path, options), options);
        }

        internal static IReadOnlyList<string> ReadSheetNames(string path, ExcelReaderOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            string extension = Path.GetExtension(path).ToLowerInvariant();
            IReadOnlyList<string> names = extension switch {
                ".xls" => LegacyXlsTabularWorkbook.ReadSheetNames(
                    path,
                    options,
                    options.CancellationToken),
                ".xlsb" => XlsbTabularWorkbook.ReadSheetNames(
                    path,
                    options,
                    options.CancellationToken),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" =>
                    ReadOpenXmlSheetNames(path, options),
                _ => throw new NotSupportedException(
                    "GetSheetNames supports .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb, and .xls workbooks.")
            };
            options.CancellationToken.ThrowIfCancellationRequested();
            return names;
        }

        private static IReadOnlyList<string> ReadOpenXmlSheetNames(
            string path,
            ExcelReaderOptions options) {
            try {
                return XlsxTabularWorkbook.ReadSheetNames(path, options);
            } catch (XlsxTabularFastPathNotSupportedException) {
                // Unusual package metadata retains the complete SDK validation path without
                // opening any worksheet data stream.
                using ExcelDocumentReader owner = ExcelDocumentReader.Open(path, options);
                IReadOnlyList<string> names = owner.GetValidatedWorksheetNames();
                ValidateUniqueSheetNames(names, options.CancellationToken);
                return names;
            }
        }

        internal static ExcelWorkbookDataReader OpenOpenXml(byte[] bytes, ExcelReaderOptions options) {
            if (CanUseNativeOpenXmlPath(options)) {
                try {
                    return CreateNativeOpenXml(XlsxTabularWorkbook.Open(bytes, options), options, () => ExcelDocumentReader.Open(bytes, options));
                } catch (XlsxTabularFastPathNotSupportedException) {
                    // Unusual package and worksheet shapes retain the complete SDK path.
                }
            }

            return CreateOpenXml(ExcelDocumentReader.Open(bytes, options), options);
        }

        internal static ExcelWorkbookDataReader WrapOpenXml(ExcelDocumentReader owner, ExcelReaderOptions options) =>
            CreateOpenXml(owner, options);

        internal static ExcelWorkbookDataReader OpenBinary(string path, ExcelReaderOptions options) =>
            string.IsNullOrWhiteSpace(options.A1Range)
                ? CreateBinary(XlsbTabularWorkbook.Open(path, options, options.CancellationToken), options)
                : OpenProjectedBinary(path, options);

        internal static ExcelWorkbookDataReader OpenBinary(byte[] bytes, ExcelReaderOptions options) =>
            string.IsNullOrWhiteSpace(options.A1Range)
                ? CreateBinary(XlsbTabularWorkbook.Open(bytes, options, options.CancellationToken), options)
                : OpenProjectedBinary(bytes, options);

        private static ExcelWorkbookDataReader OpenProjectedBinary(string path, ExcelReaderOptions options) {
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

        private static ExcelWorkbookDataReader OpenProjectedBinary(byte[] bytes, ExcelReaderOptions options) {
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

        private static ExcelLoadOptions CreateProjectedBinaryLoadOptions(ExcelReaderOptions options) =>
            new() {
                AccessMode = DocumentAccessMode.ReadOnly,
                PersistenceMode = DocumentPersistenceMode.Explicit,
                MaxInputBytes = options.MaxInputBytes,
                XlsbImportOptions = new OfficeIMO.Excel.Xlsb.XlsbImportOptions {
                    MaxPackageBytes = options.MaxInputBytes,
                    MaxWorksheets = options.MaxWorksheets,
                    MaxCells = options.MaxXlsbCells,
                    MaxLogicalRows = options.MaxXlsbLogicalRows,
                    MaxSharedStrings = options.MaxSharedStringItems,
                    MaxSharedStringItemCharacters = options.MaxSharedStringItemCharacters,
                    MaxSharedStringCharacters = options.MaxSharedStringCharacters,
                    ReportPreservedRecords = false,
                    CancellationToken = options.CancellationToken
                }
            };

        private static ExcelWorkbookDataReader WrapProjectedBinary(ExcelDocument document, ExcelReaderOptions options) {
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

        internal static ExcelWorkbookDataReader OpenLegacy(string path, ExcelReaderOptions options) {
            if (CanUseLegacyFastPath(options)) {
                try {
                    return CreateLegacy(
                        LegacyXlsTabularWorkbook.Open(path, options, options.CancellationToken),
                        options);
                } catch (LegacyXlsFastPathNotSupportedException) {
                    // Unsupported legacy variants retain the full import/projection path.
                }
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadRemainingBytes(
                stream,
                options.CancellationToken,
                options.MaxInputBytes);
            return OpenLegacyCore(bytes, path, options);
        }

        internal static ExcelWorkbookDataReader OpenLegacy(byte[] bytes, ExcelReaderOptions options) {
            if (CanUseLegacyFastPath(options)) {
                try {
                    return CreateLegacy(
                        LegacyXlsTabularWorkbook.Open(bytes, options, options.CancellationToken),
                        options);
                } catch (LegacyXlsFastPathNotSupportedException) {
                    // Unsupported legacy variants retain the full import/projection path.
                }
            }

            return OpenLegacyCore(bytes, sourcePath: null, options);
        }

        private static bool CanUseLegacyFastPath(ExcelReaderOptions options) =>
            string.IsNullOrWhiteSpace(options.A1Range)
            && options.CellValueConverter == null
            && options.UseCachedFormulaResult;

        private static ExcelWorkbookDataReader OpenLegacyCore(
            byte[] bytes,
            string? sourcePath,
            ExcelReaderOptions options) {
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
            ExcelReaderOptions options) {
            if (options.UseCachedFormulaResult) {
                return;
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<SheetSelection> sheets = SelectSheets(
                workbook.Worksheets.Select(static worksheet => worksheet.Name).ToArray(),
                options);
            (int r1, int c1, int r2, int c2)? range = string.IsNullOrWhiteSpace(options.A1Range)
                ? null
                : ExcelA1.ParseRange(options.A1Range!);
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

                    string reference = ExcelA1.CellReference(cell.Row, cell.Column);
                    throw new NotSupportedException(
                        $"Legacy XLS formula text cannot be projected for '{worksheet.Name}'!{reference}. " +
                        "UseCachedFormulaResult=false cannot return the cached value as ordinary data.");
                }
            }
        }

        internal static LegacyXlsImportOptions CreateLegacyImportOptions(ExcelReaderOptions options) =>
            new() {
                MaxInputBytes = options.MaxInputBytes > int.MaxValue
                    ? int.MaxValue
                    : checked((int)options.MaxInputBytes),
                CancellationToken = options.CancellationToken
            };

        private static ExcelWorkbookDataReader CreateLegacy(
            ExcelDocument document,
            ExcelReaderOptions options) {
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
            ExcelReaderOptions options,
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
                    (index, cancellationToken) => OpenOpenXmlSheet(
                        owner,
                        sheets[index].Name,
                        options,
                        cancellationToken),
                    lifetime,
                    options.Culture,
                    options.TypeConverter,
                    options.StrictTypedMapping,
                    options.MappingErrorValuePolicy,
                    options.CancellationToken);
            } catch {
                lifetime.Dispose();
                throw;
            }
        }

        private static ExcelWorkbookDataReader CreateNativeOpenXml(
            XlsxTabularWorkbook owner,
            ExcelReaderOptions options, Func<ExcelDocumentReader> openFallback) {
            var lifetime = new NativeOpenXmlLifetime(owner, openFallback);
            try {
                ValidateUniqueSheetNames(owner.TableNames, options.CancellationToken);
                IReadOnlyList<SheetSelection> sheets = SelectSheets(owner.TableNames, options);
                if (sheets.Count != 1) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        "Multi-result XLSX reads retain the complete Open XML SDK path.");
                }

                return new ExcelWorkbookDataReader(
                    sheets,
                    (index, cancellationToken) => lifetime.OpenSheet(sheets[index].Name, options, cancellationToken),
                    lifetime,
                    options.Culture,
                    options.TypeConverter,
                    options.StrictTypedMapping,
                    options.MappingErrorValuePolicy,
                    options.CancellationToken);
            } catch {
                lifetime.Dispose();
                throw;
            }
        }

        private static bool CanUseNativeOpenXmlPath(ExcelReaderOptions options) =>
            options.CellValueConverter == null
            && (options.InferSchema || string.IsNullOrWhiteSpace(options.A1Range));

        private static DbDataReader OpenOpenXmlSheet(
            ExcelDocumentReader owner,
            string sheetName,
            ExcelReaderOptions options,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            ExcelSheetReader sheet = owner.GetSheet(sheetName);
            return string.IsNullOrWhiteSpace(options.A1Range)
                ? (DbDataReader)sheet.ReadUsedRangeAsDataReader(
                    headersInFirstRow: options.HasHeaderRow,
                    schemaSampleRows: options.InferSchema ? options.SchemaSampleRows : 0,
                    ct: cancellationToken)
                : (DbDataReader)sheet.ReadRangeAsDataReader(
                    options.A1Range!,
                    headersInFirstRow: options.HasHeaderRow,
                    chunkRows: Math.Min(1024, options.MaxDataReaderChunkRows),
                    schemaSampleRows: options.InferSchema ? options.SchemaSampleRows : 0,
                    ct: cancellationToken);
        }

        private static ExcelWorkbookDataReader CreateBinary(
            XlsbTabularWorkbook owner,
            ExcelReaderOptions options) {
            try {
                IReadOnlyList<SheetSelection> sheets = SelectSheets(owner.TableNames, options);
                return new ExcelWorkbookDataReader(
                    sheets,
                    (index, cancellationToken) => owner.OpenTable(
                        sheets[index].Name,
                        options.HasHeaderRow,
                        options,
                        cancellationToken),
                    owner,
                    options.Culture,
                    options.TypeConverter,
                    options.StrictTypedMapping,
                    options.MappingErrorValuePolicy,
                    options.CancellationToken);
            } catch {
                owner.Dispose();
                throw;
            }
        }

        private static ExcelWorkbookDataReader CreateLegacy(
            LegacyXlsTabularWorkbook owner,
            ExcelReaderOptions options) {
            try {
                IReadOnlyList<SheetSelection> sheets = SelectSheets(owner.TableNames, options);
                return new ExcelWorkbookDataReader(
                    sheets,
                    (index, cancellationToken) => owner.OpenTable(
                        sheets[index].Name,
                        options.HasHeaderRow,
                        options,
                        cancellationToken),
                    owner,
                    options.Culture,
                    options.TypeConverter,
                    options.StrictTypedMapping,
                    options.MappingErrorValuePolicy,
                    options.CancellationToken);
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
            ExcelReaderOptions options) {
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
            return MoveNextResult(_cancellationToken);
        }

        private bool MoveNextResult(CancellationToken cancellationToken) {
            ThrowIfClosed();
            cancellationToken.ThrowIfCancellationRequested();
            if (_resultIndex + 1 >= _sheetNames.Count) {
                return false;
            }

            int nextResultIndex = _resultIndex + 1;
            DbDataReader nextReader;
            try {
                nextReader = _openSheet(nextResultIndex, cancellationToken);
            } catch {
                CloseAfterSheetOpenFailure();
                throw;
            }

            _activeReader?.Dispose();
            _currentSheetCancellationSource?.Dispose();
            _currentSheetCancellationSource = null;
            CurrentReader = nextReader;
            _resultIndex = nextResultIndex;
            return true;
        }

        /// <inheritdoc />
        public override Task<bool> NextResultAsync(CancellationToken cancellationToken) =>
            NextResultAsyncCore(cancellationToken);

        private async Task<bool> NextResultAsyncCore(CancellationToken cancellationToken) {
            ThrowIfClosed();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken.ThrowIfCancellationRequested();
            if (_resultIndex + 1 >= _sheetNames.Count) {
                return false;
            }

            if (!cancellationToken.CanBeCanceled || cancellationToken == _cancellationToken) {
                return await Task.Run(
                        () => MoveNextResult(_cancellationToken),
                        _cancellationToken)
                    .ConfigureAwait(false);
            }

            CancellationTokenSource sheetCancellationSource = _cancellationToken.CanBeCanceled
                ? CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken)
                : new CancellationTokenSource();
            CancellationTokenRegistration callerRegistration = cancellationToken.Register(
                static state => ((CancellationTokenSource)state!).Cancel(),
                sheetCancellationSource);
            DbDataReader? nextReader = null;

            try {
                int nextResultIndex = _resultIndex + 1;
                nextReader = await Task.Run(
                        () => _openSheet(nextResultIndex, sheetCancellationSource.Token),
                        sheetCancellationSource.Token)
                    .ConfigureAwait(false);

                callerRegistration.Dispose();
                cancellationToken.ThrowIfCancellationRequested();
                _cancellationToken.ThrowIfCancellationRequested();

                _activeReader?.Dispose();
                _currentSheetCancellationSource?.Dispose();
                CurrentReader = nextReader;
                _currentSheetCancellationSource = sheetCancellationSource;
                nextReader = null;
                _resultIndex = nextResultIndex;
                return true;
            } catch {
                callerRegistration.Dispose();
                nextReader?.Dispose();
                sheetCancellationSource.Dispose();
                CloseAfterSheetOpenFailure();
                throw;
            } finally {
                callerRegistration.Dispose();
            }
        }

        private CancellationTokenSource? CreateEffectiveCancellationToken(
            CancellationToken callerToken,
            out CancellationToken effectiveToken) {
            effectiveToken = callerToken;
            if (!_cancellationToken.CanBeCanceled) {
                return null;
            }
            if (!callerToken.CanBeCanceled) {
                effectiveToken = _cancellationToken;
                return null;
            }
            if (_cancellationToken == callerToken) {
                return null;
            }

            CancellationTokenSource linkedSource = CancellationTokenSource.CreateLinkedTokenSource(
                _cancellationToken,
                callerToken);
            effectiveToken = linkedSource.Token;
            return linkedSource;
        }

        private void CloseAfterSheetOpenFailure() {
            _closed = true;
            try {
                _activeReader?.Dispose();
            } catch {
                // Preserve the worksheet-open failure while attempting complete cleanup.
            }
            _currentSheetCancellationSource?.Dispose();
            _currentSheetCancellationSource = null;
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
                _activeReader?.Dispose();
            } finally {
                try {
                    _currentSheetCancellationSource?.Dispose();
                    _currentSheetCancellationSource = null;
                } finally {
                    _owner.Dispose();
                }
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

        public override object this[int ordinal] => CurrentReader[ordinal];
        /// <inheritdoc />
        public override object this[string name] => CurrentReader[name];
        /// <inheritdoc />
        public override int Depth => CurrentReader.Depth;
        /// <inheritdoc />
        public override int FieldCount => CurrentReader.FieldCount;
        /// <inheritdoc />
        public override bool HasRows => CurrentReader.HasRows;
        /// <inheritdoc />
        public override bool IsClosed => _closed || _activeReader?.IsClosed == true;
        /// <inheritdoc />
        public override int RecordsAffected => CurrentReader.RecordsAffected;
        /// <inheritdoc />
        public override int VisibleFieldCount => CurrentReader.VisibleFieldCount;
        /// <inheritdoc />
        public override bool GetBoolean(int ordinal) => CurrentReader.GetBoolean(ordinal);
        /// <inheritdoc />
        public override byte GetByte(int ordinal) => CurrentReader.GetByte(ordinal);
        /// <inheritdoc />
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
            CurrentReader.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
        /// <inheritdoc />
        public override char GetChar(int ordinal) => CurrentReader.GetChar(ordinal);
        /// <inheritdoc />
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
            CurrentReader.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
        /// <inheritdoc />
        public override string GetDataTypeName(int ordinal) => CurrentReader.GetDataTypeName(ordinal);
        /// <inheritdoc />
        public override DateTime GetDateTime(int ordinal) => CurrentReader.GetDateTime(ordinal);
        /// <inheritdoc />
        public override decimal GetDecimal(int ordinal) => CurrentReader.GetDecimal(ordinal);
        /// <inheritdoc />
        public override double GetDouble(int ordinal) => CurrentReader.GetDouble(ordinal);
        /// <inheritdoc />
        public override IEnumerator GetEnumerator() => CurrentReader.GetEnumerator();
        /// <inheritdoc />
#if NET8_0_OR_GREATER
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
        public override Type GetFieldType(int ordinal) => CurrentReader.GetFieldType(ordinal);

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
#if NET6_0_OR_GREATER
            if (destinationType == typeof(DateOnly) || destinationType == typeof(TimeOnly)) {
                object? sourceValue = IsDBNull(ordinal) ? null : GetValue(ordinal);
                return DataValueConverter.ConvertTo<T>(
                    sourceValue,
                    _culture,
                    errorValuePolicy: _mappingErrorValuePolicy)!;
            }
#endif
            if (destinationType == typeof(Guid)) return (T)(object)GetGuid(ordinal);

            object value = GetValue(ordinal);
            if (value is T typed) {
                return typed;
            }

            return (T)Convert.ChangeType(value, destinationType, _culture);
        }

        /// <inheritdoc />

        public override float GetFloat(int ordinal) => CurrentReader.GetFloat(ordinal);
        /// <inheritdoc />
        public override Guid GetGuid(int ordinal) => CurrentReader.GetGuid(ordinal);
        /// <inheritdoc />
        public override short GetInt16(int ordinal) => CurrentReader.GetInt16(ordinal);
        /// <inheritdoc />
        public override int GetInt32(int ordinal) => CurrentReader.GetInt32(ordinal);
        /// <inheritdoc />
        public override long GetInt64(int ordinal) => CurrentReader.GetInt64(ordinal);
        /// <inheritdoc />
        public override string GetName(int ordinal) => CurrentReader.GetName(ordinal);
        /// <inheritdoc />
        public override int GetOrdinal(string name) => CurrentReader.GetOrdinal(name);
        /// <inheritdoc />
        public override string GetString(int ordinal) => CurrentReader.GetString(ordinal);
        /// <inheritdoc />
        public override object GetValue(int ordinal) => CurrentReader.GetValue(ordinal);
        /// <inheritdoc />
        public override int GetValues(object[] values) => CurrentReader.GetValues(values);
        /// <inheritdoc />
        public override bool IsDBNull(int ordinal) => CurrentReader.IsDBNull(ordinal);
        /// <inheritdoc />
        public override bool Read() {
            ThrowIfClosed();
            _cancellationToken.ThrowIfCancellationRequested();
            return CurrentReader.Read();
        }
        /// <inheritdoc />
        public override async Task<bool> ReadAsync(CancellationToken cancellationToken) {
            ThrowIfClosed();
            CancellationTokenSource? linkedSource = CreateEffectiveCancellationToken(
                cancellationToken,
                out CancellationToken effectiveToken);
            try {
                effectiveToken.ThrowIfCancellationRequested();
                return await GetCurrentReader(effectiveToken).ReadAsync(effectiveToken).ConfigureAwait(false);
            } finally {
                linkedSource?.Dispose();
            }
        }
        /// <inheritdoc />
        public override DataTable? GetSchemaTable() => CurrentReader.GetSchemaTable();

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
