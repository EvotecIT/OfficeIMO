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
        private const int BrtFont = 43;
        private const int BrtFmt = 44;
        private const int BrtFill = 45;
        private const int BrtBorder = 46;
        private const int BrtXf = 47;
        private const int BrtBeginFmts = 615;
        private const int BrtEndFmts = 616;
        private const int BrtBeginCellXfs = 617;
        private const int BrtEndCellXfs = 618;
        private const int BrtBeginCellStyleXfs = 626;
        private const int BrtEndCellStyleXfs = 627;
        private const int MaxStyleItems = 65_536;
        private const int MaxFonts = 0xFFD3;
        private const int MaxCellFormats = 0xFF96;
        private const string WorksheetRelationshipSuffix = "/worksheet";
        private const string ChartSheetRelationshipSuffix = "/chartsheet";
        private const string DialogSheetRelationshipSuffix = "/dialogsheet";
        private const string MacroSheetRelationshipSuffix = "/macrosheet";
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

            XlsbPooledPartStream part = _parts.ReadSeekablePart(sheet.PartName, cancellationToken);
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
            XlsbPackageRelationship? relationship = XlsbPackagePartReader.GetOptionalSingletonRelationship(
                relationships,
                SharedStringsRelationshipSuffix,
                "shared-string");
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
            XlsbPackageRelationship? relationship = XlsbPackagePartReader.GetOptionalSingletonRelationship(
                relationships,
                StylesRelationshipSuffix,
                "styles");
            if (relationship == null) {
                return new[] { false };
            }

            string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
            byte[] bytes = _parts.ReadPart(partName, _cancellationToken);
            var records = new XlsbRecordSliceReader(bytes, _limits.MaxRecordBytes, _recordBudget);
            var customFormats = new Dictionary<ushort, string>();
            var cellStyleFormats = new List<XlsbTabularCellFormatReference>();
            var cellFormats = new List<XlsbTabularCellFormatReference>();
            var seenCollectionBegins = new HashSet<int>();
            int activeCollectionEnd = 0;
            int declaredCollectionCount = 0;
            int actualCollectionCount = 0;
            int fontCount = 0;
            int fillCount = 0;
            int borderCount = 0;
            string? activeCollectionName = null;
            bool beganStyleSheet = false;
            bool endedStyleSheet = false;
            bool sawCellFormats = false;
            while (records.TryRead(out XlsbRecordSlice record)) {
                CheckCancellation();
                if (endedStyleSheet) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains records after BrtEndStyleSheet.");
                }
                if (IsSupportedStyleCollectionBegin(record.Type)
                    && !seenCollectionBegins.Add(record.Type)) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains more than one {GetStyleCollectionName(record.Type)} collection.");
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
                        BeginStyleCollection(
                            partName,
                            "number-format",
                            BrtEndFmts,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        break;
                    case BrtEndFmts:
                        EndStyleCollection(
                            partName,
                            BrtEndFmts,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        break;
                    case BrtBeginCellXfs:
                        BeginStyleCollection(
                            partName,
                            "cell-format",
                            BrtEndCellXfs,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        sawCellFormats = true;
                        break;
                    case BrtEndCellXfs:
                        EndStyleCollection(
                            partName,
                            BrtEndCellXfs,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        break;
                    case BrtBeginFills:
                        BeginStyleCollection(
                            partName,
                            "fill",
                            BrtEndFills,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        break;
                    case BrtEndFills:
                        EndStyleCollection(
                            partName,
                            BrtEndFills,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        fillCount = checked(fillCount + actualCollectionCount);
                        break;
                    case BrtBeginFonts:
                        BeginStyleCollection(
                            partName,
                            "font",
                            BrtEndFonts,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        break;
                    case BrtEndFonts:
                        EndStyleCollection(
                            partName,
                            BrtEndFonts,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        fontCount = checked(fontCount + actualCollectionCount);
                        break;
                    case BrtBeginBorders:
                        BeginStyleCollection(
                            partName,
                            "border",
                            BrtEndBorders,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        break;
                    case BrtEndBorders:
                        EndStyleCollection(
                            partName,
                            BrtEndBorders,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        borderCount = checked(borderCount + actualCollectionCount);
                        break;
                    case BrtBeginCellStyleXfs:
                        BeginStyleCollection(
                            partName,
                            "cell-style-format",
                            BrtEndCellStyleXfs,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            ref declaredCollectionCount,
                            ref actualCollectionCount);
                        break;
                    case BrtEndCellStyleXfs:
                        EndStyleCollection(
                            partName,
                            BrtEndCellStyleXfs,
                            record,
                            ref activeCollectionEnd,
                            ref activeCollectionName,
                            declaredCollectionCount,
                            actualCollectionCount);
                        break;
                    case BrtFmt when activeCollectionEnd == BrtEndFmts: {
                        var cursor = record.CreateCursor();
                        ushort numberFormatId = cursor.ReadUInt16();
                        string formatCode = cursor.ReadWideString(
                            Math.Min(_limits.MaxStringCharacters, 255));
                        if (cursor.Remaining != 0 || formatCode.Length == 0) {
                            throw new InvalidDataException(
                                $"The BrtFmt record at offset {record.RecordOffset} is malformed.");
                        }
                        if (customFormats.ContainsKey(numberFormatId)) {
                            throw new InvalidDataException(
                                $"The XLSB styles part '{partName}' contains duplicate custom number format {numberFormatId}.");
                        }
                        customFormats.Add(numberFormatId, formatCode);
                        actualCollectionCount++;
                        break;
                    }
                    case BrtXf when activeCollectionEnd == BrtEndCellXfs: {
                        XlsbTabularCellFormatReference format = ReadCellFormatReference(record);
                        cellFormats.Add(format);
                        actualCollectionCount++;
                        break;
                    }
                    case BrtXf when activeCollectionEnd == BrtEndCellStyleXfs:
                        cellStyleFormats.Add(ReadCellFormatReference(record));
                        actualCollectionCount++;
                        break;
                    case BrtFont when activeCollectionEnd == BrtEndFonts:
                        ValidateFontPayload(record);
                        actualCollectionCount++;
                        break;
                    case BrtFill when activeCollectionEnd == BrtEndFills:
                        ValidateFillPayload(record);
                        actualCollectionCount++;
                        break;
                    case BrtBorder when activeCollectionEnd == BrtEndBorders:
                        ValidateBorderPayload(record);
                        actualCollectionCount++;
                        break;
                }
            }

            if (!beganStyleSheet || !endedStyleSheet || activeCollectionEnd != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' is missing required stylesheet or collection boundary records.");
            }
            if (!sawCellFormats || cellFormats.Count == 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' is missing the required non-empty cell-format collection.");
            }
            if (fontCount == 0 || fillCount == 0 || borderCount == 0 || cellStyleFormats.Count == 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' is missing one or more required formatting collections.");
            }
            ValidateCellFormatReferences(
                partName,
                cellStyleFormats,
                cellFormats,
                fontCount,
                fillCount,
                borderCount,
                customFormats);

            return cellFormats
                .Select(format =>
                    ExcelBuiltInNumberFormats.IsDate(format.NumberFormatId)
                    || (customFormats.TryGetValue(format.NumberFormatId, out string? code)
                        && ExcelNumberFormatClassifier.LooksLikeDateFormat(code)))
                .ToArray();
        }

        private static bool IsSupportedStyleCollectionBegin(int recordType) =>
            recordType == BrtBeginFmts
            || recordType == BrtBeginFonts
            || recordType == BrtBeginFills
            || recordType == BrtBeginBorders
            || recordType == BrtBeginCellStyleXfs
            || recordType == BrtBeginCellXfs;

        private static string GetStyleCollectionName(int recordType) {
            switch (recordType) {
                case BrtBeginFmts:
                    return "number-format";
                case BrtBeginFonts:
                    return "font";
                case BrtBeginFills:
                    return "fill";
                case BrtBeginBorders:
                    return "border";
                case BrtBeginCellStyleXfs:
                    return "cell-style-format";
                case BrtBeginCellXfs:
                    return "cell-format";
                default:
                    throw new ArgumentOutOfRangeException(nameof(recordType));
            }
        }

        private void ValidateFontPayload(XlsbRecordSlice record) {
            try {
                var cursor = record.CreateCursor();
                cursor.Skip(12);
                cursor.Skip(8);
                cursor.ReadByte();
                string name = cursor.ReadWideString(Math.Min(_limits.MaxStringCharacters, 31));
                if (cursor.Remaining != 0 || name.Length == 0) {
                    throw new InvalidDataException(
                        $"The BrtFont record at offset {record.RecordOffset} is malformed.");
                }
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException(
                    $"The BrtFont record at offset {record.RecordOffset} is truncated.",
                    exception);
            }
        }

        private static void ValidateFillPayload(XlsbRecordSlice record) {
            try {
                var cursor = record.CreateCursor();
                cursor.ReadUInt32();
                cursor.Skip(8);
                cursor.Skip(8);
                cursor.ReadInt32();
                cursor.Skip(8 * 5);
                uint gradientStopCount = cursor.ReadUInt32();
                if (gradientStopCount > MaxStyleItems) {
                    throw new InvalidDataException(
                        $"The BrtFill record at offset {record.RecordOffset} declares too many gradient stops.");
                }
                for (uint index = 0; index < gradientStopCount; index++) {
                    cursor.ReadDouble();
                    cursor.Skip(8);
                }
                if (cursor.Remaining != 0) {
                    throw new InvalidDataException(
                        $"The BrtFill record at offset {record.RecordOffset} has unexpected trailing data.");
                }
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException(
                    $"The BrtFill record at offset {record.RecordOffset} is truncated.",
                    exception);
            }
        }

        private static void ValidateBorderPayload(XlsbRecordSlice record) {
            try {
                var cursor = record.CreateCursor();
                cursor.ReadByte();
                for (int side = 0; side < 5; side++) {
                    cursor.ReadByte();
                    cursor.Skip(1);
                    cursor.Skip(8);
                }
                if (cursor.Remaining != 0) {
                    throw new InvalidDataException(
                        $"The BrtBorder record at offset {record.RecordOffset} has unexpected trailing data.");
                }
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException(
                    $"The BrtBorder record at offset {record.RecordOffset} is truncated.",
                    exception);
            }
        }

        private static XlsbTabularCellFormatReference ReadCellFormatReference(
            XlsbRecordSlice record) {
            if (record.Size != 16) {
                throw new InvalidDataException(
                    $"The BrtXf record at offset {record.RecordOffset} has invalid payload length {record.Size}.");
            }

            var cursor = record.CreateCursor();
            return new XlsbTabularCellFormatReference(
                cursor.ReadUInt16(),
                cursor.ReadUInt16(),
                cursor.ReadUInt16(),
                cursor.ReadUInt16(),
                cursor.ReadUInt16());
        }

        private static void ValidateCellFormatReferences(
            string partName,
            IReadOnlyList<XlsbTabularCellFormatReference> cellStyleFormats,
            IReadOnlyList<XlsbTabularCellFormatReference> cellFormats,
            int fontCount,
            int fillCount,
            int borderCount,
            IReadOnlyDictionary<ushort, string> customFormats) {
            foreach (XlsbTabularCellFormatReference format in cellStyleFormats.Concat(cellFormats)) {
                if (format.FontId >= fontCount
                    || format.FillId >= fillCount
                    || format.BorderId >= borderCount) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains a cell format with an out-of-range font, fill, or border reference.");
                }
                if (format.NumberFormatId >= ExcelBuiltInNumberFormats.FirstCustomId
                    && !customFormats.ContainsKey(format.NumberFormatId)) {
                    throw new InvalidDataException(
                        $"The XLSB styles part '{partName}' contains a cell format that references missing custom number format {format.NumberFormatId}.");
                }
            }

            if (cellStyleFormats.Any(format => format.ParentFormatId != ushort.MaxValue)
                || cellFormats.Any(format => format.ParentFormatId >= cellStyleFormats.Count)) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' contains an invalid parent cell-style reference.");
            }
        }

        private readonly struct XlsbTabularCellFormatReference {
            internal XlsbTabularCellFormatReference(
                ushort parentFormatId,
                ushort numberFormatId,
                ushort fontId,
                ushort fillId,
                ushort borderId) {
                ParentFormatId = parentFormatId;
                NumberFormatId = numberFormatId;
                FontId = fontId;
                FillId = fillId;
                BorderId = borderId;
            }

            internal ushort ParentFormatId { get; }

            internal ushort NumberFormatId { get; }

            internal ushort FontId { get; }

            internal ushort FillId { get; }

            internal ushort BorderId { get; }
        }

        private static void BeginStyleCollection(
            string partName,
            string collectionName,
            int expectedEndRecord,
            XlsbRecordSlice record,
            ref int activeCollectionEnd,
            ref string? activeCollectionName,
            ref int declaredCollectionCount,
            ref int actualCollectionCount) {
            if (activeCollectionEnd != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' begins a style collection before the active collection ends.");
            }
            if (record.Size != sizeof(uint)) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' contains an invalid {collectionName} collection header.");
            }

            uint declared = record.CreateCursor().ReadUInt32();
            int maximum = collectionName == "font"
                ? MaxFonts
                : collectionName == "cell-format"
                    ? MaxCellFormats
                    : MaxStyleItems;
            if (declared > maximum) {
                throw new InvalidDataException(
                    $"The XLSB {collectionName} collection declares {declared} items, exceeding the supported limit of {maximum}.");
            }

            activeCollectionEnd = expectedEndRecord;
            activeCollectionName = collectionName;
            declaredCollectionCount = checked((int)declared);
            actualCollectionCount = 0;
        }

        private static void EndStyleCollection(
            string partName,
            int endRecord,
            XlsbRecordSlice record,
            ref int activeCollectionEnd,
            ref string? activeCollectionName,
            int declaredCollectionCount,
            int actualCollectionCount) {
            if (activeCollectionEnd != endRecord || record.Size != 0) {
                throw new InvalidDataException(
                    $"The XLSB styles part '{partName}' contains an invalid style collection end record.");
            }
            if (actualCollectionCount != declaredCollectionCount) {
                throw new InvalidDataException(
                    $"The XLSB {activeCollectionName} collection declares {declaredCollectionCount} items but contains {actualCollectionCount} item records.");
            }

            activeCollectionEnd = 0;
            activeCollectionName = null;
        }

        private List<XlsbTabularSheet> ResolveWorksheets(
            string workbookPartName,
            IReadOnlyDictionary<string, XlsbPackageRelationship> relationships,
            IReadOnlyList<XlsbBundleSheet> bundleSheets) {
            var sheets = new List<XlsbTabularSheet>(bundleSheets.Count);
            var worksheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XlsbBundleSheet bundle in bundleSheets) {
                if (!relationships.TryGetValue(
                        bundle.RelationshipId,
                        out XlsbPackageRelationship? relationship)) {
                    throw new InvalidDataException(
                        $"The XLSB worksheet '{bundle.Name}' references missing relationship '{bundle.RelationshipId}'.");
                }
                if (relationship.IsExternal) {
                    throw new InvalidDataException(
                        $"The XLSB worksheet '{bundle.Name}' references external relationship '{bundle.RelationshipId}'.");
                }
                if (!relationship.Type.EndsWith(WorksheetRelationshipSuffix, StringComparison.Ordinal)) {
                    if (IsSupportedNonWorksheetSheetRelationship(relationship.Type)) {
                        continue;
                    }

                    throw new InvalidDataException(
                        $"The XLSB sheet '{bundle.Name}' references unrelated internal relationship '{bundle.RelationshipId}'.");
                }
                if (!worksheetNames.Add(bundle.Name)) {
                    throw new InvalidDataException(
                        $"The XLSB workbook contains duplicate worksheet name '{bundle.Name}'.");
                }

                string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, relationship.Target);
                sheets.Add(new XlsbTabularSheet(bundle.Name, partName));
            }

            return sheets;
        }

        private static bool IsSupportedNonWorksheetSheetRelationship(string relationshipType) =>
            relationshipType.EndsWith(ChartSheetRelationshipSuffix, StringComparison.Ordinal)
            || relationshipType.EndsWith(DialogSheetRelationshipSuffix, StringComparison.Ordinal)
            || relationshipType.EndsWith(MacroSheetRelationshipSuffix, StringComparison.Ordinal);

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
