using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using System.Data.Common;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Owns the bounded package state needed by forward-only XLSB table readers. Unlike the
    /// editable importer, it does not build an intermediate workbook or Open XML projection.
    /// </summary>
    internal sealed class XlsbTabularWorkbook : IDisposable {
        private const int BrtBeginBook = 131;
        private const int BrtEndBook = 132;
        private const int BrtWbProp = 153;
        private const int BrtBundleSh = 156;
        private const int BrtSstItem = 19;
        private const int BrtBeginSst = 159;
        private const int BrtEndSst = 160;
        private const int BrtBeginStyleSheet = 278;
        private const int BrtEndStyleSheet = 279;
        private const int BrtBeginFills = 603;
        private const int BrtEndFills = 604;
        private const int BrtBeginFonts = 611;
        private const int BrtEndFonts = 612;
        private const int BrtBeginBorders = 613;
        private const int BrtEndBorders = 614;
        private const int BrtFmt = 44;
        private const int BrtXf = 47;
        private const int BrtBeginFmts = 615;
        private const int BrtEndFmts = 616;
        private const int BrtBeginCellXfs = 617;
        private const int BrtEndCellXfs = 618;
        private const int BrtBeginCellStyleXfs = 626;
        private const int BrtEndCellStyleXfs = 627;
        private const string WorksheetRelationshipSuffix = "/worksheet";
        private const string SharedStringsRelationshipSuffix = "/sharedStrings";
        private const string StylesRelationshipSuffix = "/styles";

        private readonly Stream _packageStream;
        private readonly ZipArchive _archive;
        private readonly XlsbPackagePartReader _parts;
        private readonly XlsbImportOptions _limits;
        private readonly XlsbRecordReadBudget _recordBudget;
        private readonly XlsbCellReadBudget _cellBudget;
        private readonly IReadOnlyList<string> _sharedStrings;
        private readonly int _maxSharedStringItemCharacters;
        private readonly long _maxSharedStringCharacters;
        private readonly bool[] _dateStyles;
        private readonly List<XlsbTabularSheet> _sheets;
        private readonly string[] _tableNames;
        private readonly CancellationToken _cancellationToken;
        private int _recordsSinceCancellationCheck;
        private bool _disposed;

        private XlsbTabularWorkbook(
            Stream packageStream,
            ZipArchive archive,
            string workbookPartName,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken) {
            _packageStream = packageStream ?? throw new ArgumentNullException(nameof(packageStream));
            _archive = archive ?? throw new ArgumentNullException(nameof(archive));
            _cancellationToken = cancellationToken;
            _maxSharedStringItemCharacters = readOptions.MaxSharedStringItemCharacters;
            _maxSharedStringCharacters = readOptions.MaxSharedStringCharacters;
            _limits = new XlsbImportOptions {
                MaxPackageBytes = readOptions.MaxInputBytes,
                MaxSharedStrings = readOptions.MaxSharedStringItems,
                MaxCells = readOptions.MaxXlsbCells,
                ReportPreservedRecords = false
            };
            _limits.Validate();
            _parts = new XlsbPackagePartReader(_archive, _limits);
            _recordBudget = new XlsbRecordReadBudget(_limits.MaxRecordCount);
            _cellBudget = new XlsbCellReadBudget(_limits.MaxCells);

            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships =
                _parts.ReadRelationships(workbookPartName, cancellationToken);
            var bundleSheets = new List<XlsbBundleSheet>();
            bool uses1904DateSystem = false;
            ParseWorkbookPart(
                _parts.ReadPart(workbookPartName, cancellationToken),
                bundleSheets,
                ref uses1904DateSystem);
            Uses1904DateSystem = uses1904DateSystem;
            _sharedStrings = ReadSharedStrings(workbookPartName, relationships);
            _dateStyles = ReadDateStyles(workbookPartName, relationships);
            _sheets = ResolveWorksheets(workbookPartName, relationships, bundleSheets);
            if (_sheets.Count == 0) {
                throw new InvalidDataException("The XLSB workbook contains no readable worksheets.");
            }
            _tableNames = _sheets.Select(static sheet => sheet.Name).ToArray();
        }

        internal bool Uses1904DateSystem { get; }

        internal IReadOnlyList<string> TableNames => _tableNames;

        internal static XlsbTabularWorkbook Open(
            string path,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }

            if (readOptions == null) {
                throw new ArgumentNullException(nameof(readOptions));
            }

            cancellationToken.ThrowIfCancellationRequested();
            var stream = File.OpenRead(path);
            try {
                if (stream.Length > readOptions.MaxInputBytes) {
                    throw new InvalidDataException(
                        $"The XLSB package contains {stream.Length} bytes, exceeding the configured limit of {readOptions.MaxInputBytes} bytes.");
                }

                return OpenOwnedStream(stream, readOptions, cancellationToken);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        internal static XlsbTabularWorkbook Open(
            Stream stream,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken = default) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            if (!stream.CanRead) {
                throw new ArgumentException("Stream must be readable.", nameof(stream));
            }

            if (readOptions == null) {
                throw new ArgumentNullException(nameof(readOptions));
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = OfficeStreamReader.ReadAllBytes(
                stream,
                cancellationToken,
                readOptions.MaxInputBytes);
            return OpenOwnedStream(
                new MemoryStream(bytes, writable: false),
                readOptions,
                cancellationToken);
        }

        internal static XlsbTabularWorkbook Open(
            byte[] bytes,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken = default) {
            if (bytes == null) {
                throw new ArgumentNullException(nameof(bytes));
            }
            if (readOptions == null) {
                throw new ArgumentNullException(nameof(readOptions));
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (bytes.LongLength > readOptions.MaxInputBytes) {
                throw new InvalidDataException(
                    $"The XLSB package contains {bytes.LongLength} bytes, exceeding the configured limit of {readOptions.MaxInputBytes} bytes.");
            }

            return OpenOwnedStream(
                new MemoryStream(bytes, writable: false),
                readOptions,
                cancellationToken);
        }

        internal DbDataReader OpenTable(
            string tableName,
            bool hasHeaderRow,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken = default) {
            ThrowIfDisposed();
            XlsbTabularSheet? sheet = _sheets.FirstOrDefault(
                candidate => string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null) {
                throw new KeyNotFoundException($"Table '{tableName}' was not found.");
            }

            Stream part = new MemoryStream(
                _parts.ReadPart(sheet.PartName, cancellationToken),
                writable: false);
            try {
                return new XlsbTabularDataReader(
                    part,
                    _sharedStrings,
                    _dateStyles,
                    Uses1904DateSystem,
                    hasHeaderRow,
                    readOptions,
                    _limits,
                    _recordBudget,
                    _cellBudget,
                    cancellationToken);
            } catch {
                part.Dispose();
                throw;
            }
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            _disposed = true;
            _archive.Dispose();
            _packageStream.Dispose();
        }

        private void ParseWorkbookPart(
            byte[] bytes,
            List<XlsbBundleSheet> sheets,
            ref bool uses1904DateSystem) {
            var records = new XlsbRecordSliceReader(bytes, _limits.MaxRecordBytes, _recordBudget);
            bool began = false;
            bool ended = false;
            while (records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (!began) {
                    if (record.Type != BrtBeginBook || record.RecordOffset != 0 || record.Size != 0) {
                        throw new InvalidDataException(
                            "The XLSB workbook part is missing its initial BrtBeginBook boundary.");
                    }
                    began = true;
                    continue;
                }
                if (ended) {
                    throw new InvalidDataException(
                        "The XLSB workbook part contains records after BrtEndBook.");
                }
                if (record.Type == BrtBeginBook) {
                    throw new InvalidDataException(
                        "The XLSB workbook part contains more than one BrtBeginBook boundary.");
                }
                if (record.Type == BrtEndBook) {
                    if (record.Size != 0) {
                        throw new InvalidDataException(
                            "The XLSB workbook part contains an invalid BrtEndBook boundary.");
                    }
                    ended = true;
                    continue;
                }
                if (record.Type == BrtWbProp) {
                    var cursor = record.CreateCursor();
                    uses1904DateSystem = (cursor.ReadUInt32() & 0x01U) != 0;
                    continue;
                }

                if (record.Type != BrtBundleSh) {
                    continue;
                }

                var sheetCursor = record.CreateCursor();
                sheetCursor.ReadUInt32();
                sheetCursor.ReadUInt32();
                string relationshipId = sheetCursor.ReadWideString(_limits.MaxStringCharacters);
                string name = sheetCursor.ReadWideString(_limits.MaxStringCharacters);
                if (string.IsNullOrWhiteSpace(relationshipId) || string.IsNullOrWhiteSpace(name)) {
                    throw new InvalidDataException(
                        $"The BrtBundleSh record at offset {record.RecordOffset} does not contain a worksheet name and relationship id.");
                }

                sheets.Add(new XlsbBundleSheet(name, relationshipId));
            }

            if (!began || !ended) {
                throw new InvalidDataException(
                    "The XLSB workbook part is missing its BrtBeginBook/BrtEndBook boundaries.");
            }
            if (sheets.Count > _limits.MaxWorksheets) {
                throw new InvalidDataException(
                    $"The XLSB workbook contains {sheets.Count} worksheets, exceeding the configured limit of {_limits.MaxWorksheets}.");
            }
        }

        private IReadOnlyList<string> ReadSharedStrings(
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships) {
            XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(candidate =>
                !candidate.IsExternal
                && candidate.Type.EndsWith(SharedStringsRelationshipSuffix, StringComparison.Ordinal));
            if (relationship == null) {
                return Array.Empty<string>();
            }

            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            byte[] bytes = _parts.ReadPart(partName, _cancellationToken);
            var records = new XlsbRecordSliceReader(bytes, _limits.MaxRecordBytes, _recordBudget);
            var values = new List<string>();
            long totalCharacters = 0;
            bool began = false;
            bool ended = false;
            uint declaredUniqueCount = 0;
            while (records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (record.Type == BrtBeginSst) {
                    if (began || ended || record.RecordOffset != 0 || record.Size < 8) {
                        throw new InvalidDataException(
                            $"The BrtBeginSst record in '{partName}' is truncated or misplaced.");
                    }
                    var header = record.CreateCursor();
                    uint declaredTotalCount = header.ReadUInt32();
                    declaredUniqueCount = header.ReadUInt32();
                    if (declaredTotalCount < declaredUniqueCount) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string part '{partName}' declares fewer total strings than unique strings.");
                    }
                    if (declaredUniqueCount > _limits.MaxSharedStrings) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string table declares {declaredUniqueCount} unique items, exceeding the configured limit of {_limits.MaxSharedStrings} items.");
                    }
                    began = true;
                } else if (record.Type == BrtEndSst) {
                    if (!began || ended || record.Size != 0) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string part '{partName}' contains an invalid BrtEndSst boundary.");
                    }
                    ended = true;
                } else if (record.Type == BrtSstItem) {
                    if (!began || ended) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string part '{partName}' contains an item outside its boundary records.");
                    }
                    if (values.Count >= _limits.MaxSharedStrings) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string table exceeds the configured limit of {_limits.MaxSharedStrings} items.");
                    }

                    var cursor = record.CreateCursor();
                    cursor.ReadByte();
                    string value = cursor.ReadWideString(_maxSharedStringItemCharacters);
                    if (totalCharacters > _maxSharedStringCharacters - value.Length) {
                        throw new InvalidDataException(
                            $"The XLSB shared-string table exceeds the configured aggregate limit of {_maxSharedStringCharacters} characters.");
                    }

                    totalCharacters += value.Length;
                    values.Add(value);
                } else if (!began || ended) {
                    throw new InvalidDataException(
                        $"The XLSB shared-string part '{partName}' contains records outside its boundary records.");
                }
            }

            if (!began || !ended) {
                throw new InvalidDataException(
                    $"The XLSB shared-string part '{partName}' is missing its boundary records.");
            }
            if ((uint)values.Count != declaredUniqueCount) {
                throw new InvalidDataException(
                    $"The XLSB shared-string part '{partName}' declares {declaredUniqueCount} unique items but contains {values.Count}.");
            }

            return values;
        }

        private bool[] ReadDateStyles(
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships) {
            XlsbPackageRelationship? relationship = relationships.Values.FirstOrDefault(candidate =>
                !candidate.IsExternal
                && candidate.Type.EndsWith(StylesRelationshipSuffix, StringComparison.Ordinal));
            if (relationship == null) {
                return new[] { false };
            }

            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            byte[] bytes = _parts.ReadPart(partName, _cancellationToken);
            var records = new XlsbRecordSliceReader(bytes, _limits.MaxRecordBytes, _recordBudget);
            var customFormats = new Dictionary<ushort, string>();
            var dateStyles = new List<bool>();
            int activeCollectionEnd = 0;
            bool beganStyleSheet = false;
            bool endedStyleSheet = false;
            bool inFormats = false;
            bool inCellFormats = false;
            int declaredCellFormatCount = -1;
            int actualCellFormatCount = 0;
            while (records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (endedStyleSheet) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains records after BrtEndStyleSheet.");
                }

                switch (record.Type) {
                    case BrtBeginStyleSheet:
                        if (beganStyleSheet || record.RecordOffset != 0 || record.Size != 0) {
                            throw new InvalidDataException(
                                $"The XLSB styles part '{partName}' contains an invalid BrtBeginStyleSheet record.");
                        }
                        beganStyleSheet = true;
                        break;
                    case BrtEndStyleSheet:
                        if (!beganStyleSheet || activeCollectionEnd != 0 || record.Size != 0) {
                            throw new InvalidDataException(
                                $"The XLSB styles part '{partName}' contains an invalid BrtEndStyleSheet record.");
                        }
                        endedStyleSheet = true;
                        break;
                    case BrtBeginFmts:
                        BeginStyleCollection(partName, BrtEndFmts, ref activeCollectionEnd);
                        inFormats = true;
                        break;
                    case BrtEndFmts:
                        EndStyleCollection(partName, BrtEndFmts, record, ref activeCollectionEnd);
                        inFormats = false;
                        break;
                    case BrtBeginCellXfs:
                        BeginStyleCollection(partName, BrtEndCellXfs, ref activeCollectionEnd);
                        if (record.Size != sizeof(uint)) {
                            throw new InvalidDataException(
                                $"The XLSB styles part '{partName}' contains an invalid cell-format collection header.");
                        }
                        uint declaredCellFormats = record.CreateCursor().ReadUInt32();
                        if (declaredCellFormats > int.MaxValue) {
                            throw new InvalidDataException(
                                $"The XLSB cell-format collection declares {declaredCellFormats} items, exceeding the supported limit of {int.MaxValue}.");
                        }
                        declaredCellFormatCount = checked((int)declaredCellFormats);
                        actualCellFormatCount = 0;
                        inCellFormats = true;
                        break;
                    case BrtEndCellXfs:
                        if (actualCellFormatCount != declaredCellFormatCount) {
                            throw new InvalidDataException(
                                $"The XLSB cell-format collection declares {declaredCellFormatCount} items but contains {actualCellFormatCount} item records.");
                        }
                        EndStyleCollection(partName, BrtEndCellXfs, record, ref activeCollectionEnd);
                        inCellFormats = false;
                        break;
                    case BrtBeginFills:
                        BeginStyleCollection(partName, BrtEndFills, ref activeCollectionEnd);
                        break;
                    case BrtEndFills:
                        EndStyleCollection(partName, BrtEndFills, record, ref activeCollectionEnd);
                        break;
                    case BrtBeginFonts:
                        BeginStyleCollection(partName, BrtEndFonts, ref activeCollectionEnd);
                        break;
                    case BrtEndFonts:
                        EndStyleCollection(partName, BrtEndFonts, record, ref activeCollectionEnd);
                        break;
                    case BrtBeginBorders:
                        BeginStyleCollection(partName, BrtEndBorders, ref activeCollectionEnd);
                        break;
                    case BrtEndBorders:
                        EndStyleCollection(partName, BrtEndBorders, record, ref activeCollectionEnd);
                        break;
                    case BrtBeginCellStyleXfs:
                        BeginStyleCollection(partName, BrtEndCellStyleXfs, ref activeCollectionEnd);
                        break;
                    case BrtEndCellStyleXfs:
                        EndStyleCollection(partName, BrtEndCellStyleXfs, record, ref activeCollectionEnd);
                        break;
                    case BrtFmt when inFormats: {
                        var cursor = record.CreateCursor();
                        ushort numberFormatId = cursor.ReadUInt16();
                        customFormats[numberFormatId] = cursor.ReadWideString(
                            Math.Min(_limits.MaxStringCharacters, 255));
                        break;
                    }
                    case BrtXf when inCellFormats: {
                        var cursor = record.CreateCursor();
                        cursor.ReadUInt16();
                        ushort numberFormatId = cursor.ReadUInt16();
                        bool isDate = ExcelBuiltInNumberFormats.IsDate(numberFormatId)
                            || (customFormats.TryGetValue(numberFormatId, out string? code)
                                && ExcelNumberFormatClassifier.LooksLikeDateFormat(code));
                        dateStyles.Add(isDate);
                        actualCellFormatCount++;
                        break;
                    }
                }
            }

            if (!beganStyleSheet || !endedStyleSheet || activeCollectionEnd != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' is missing required stylesheet or collection boundary records.");
            }

            return dateStyles.Count == 0 ? new[] { false } : dateStyles.ToArray();
        }

        private static void BeginStyleCollection(
            string partName,
            int expectedEndRecord,
            ref int activeCollectionEnd) {
            if (activeCollectionEnd != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' begins a style collection before the active collection ends.");
            }

            activeCollectionEnd = expectedEndRecord;
        }

        private static void EndStyleCollection(
            string partName,
            int endRecord,
            XlsbRecordSlice record,
            ref int activeCollectionEnd) {
            if (activeCollectionEnd != endRecord || record.Size != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' contains an invalid style collection end record.");
            }

            activeCollectionEnd = 0;
        }

        private List<XlsbTabularSheet> ResolveWorksheets(
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships,
            IReadOnlyList<XlsbBundleSheet> bundleSheets) {
            var sheets = new List<XlsbTabularSheet>(bundleSheets.Count);
            foreach (XlsbBundleSheet bundle in bundleSheets) {
                if (!relationships.TryGetValue(bundle.RelationshipId, out XlsbPackageRelationship? relationship)
                    || relationship.IsExternal
                    || !relationship.Type.EndsWith(WorksheetRelationshipSuffix, StringComparison.Ordinal)) {
                    continue;
                }

                string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
                sheets.Add(new XlsbTabularSheet(bundle.Name, partName));
            }

            return sheets;
        }

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(XlsbTabularWorkbook));
            }
        }

        private static XlsbTabularWorkbook OpenOwnedStream(
            Stream packageStream,
            ExcelReadOptions readOptions,
            CancellationToken cancellationToken) {
            ZipArchive? archive = null;
            try {
                archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);
                if (!XlsbPackageDetector.TryFindWorkbookPart(archive, out string? workbookPartName)
                    || string.IsNullOrWhiteSpace(workbookPartName)) {
                    throw new InvalidDataException("The package does not contain a canonical XLSB workbook part.");
                }

                return new XlsbTabularWorkbook(
                    packageStream,
                    archive,
                    workbookPartName!,
                    readOptions,
                    cancellationToken);
            } catch {
                archive?.Dispose();
                packageStream.Dispose();
                throw;
            }
        }

        private void CheckCancellation() {
            if (!_cancellationToken.CanBeCanceled) {
                return;
            }

            _recordsSinceCancellationCheck++;
            if ((_recordsSinceCancellationCheck & 1023) == 0) {
                _cancellationToken.ThrowIfCancellationRequested();
            }
        }

        private sealed class XlsbBundleSheet {
            internal XlsbBundleSheet(string name, string relationshipId) {
                Name = name;
                RelationshipId = relationshipId;
            }

            internal string Name { get; }

            internal string RelationshipId { get; }
        }

        private sealed class XlsbTabularSheet {
            internal XlsbTabularSheet(string name, string partName) {
                Name = name;
                PartName = partName;
            }

            internal string Name { get; }

            internal string PartName { get; }
        }
    }
}
