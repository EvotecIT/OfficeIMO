using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Data.Common;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Read {
    /// <summary>
    /// Read-only BIFF5/BIFF8 workbook metadata and BIFF8 worksheet data used by the tabular
    /// XLS fast path. The workbook stream is extracted once and worksheet cells are decoded
    /// directly from record slices.
    /// </summary>
    internal sealed class LegacyXlsTabularWorkbook : IDisposable {
        private const ushort Biff5Version = 0x0500;
        private const ushort Biff8Version = 0x0600;
        private readonly LegacyBiffSource _workbookStream;
        private readonly IReadOnlyList<SheetInfo> _sheets;
        private readonly IReadOnlyList<string> _tableNames;
        private readonly IReadOnlyList<string> _sharedStrings;
        private readonly bool[] _dateStyles;
        private readonly bool _uses1904DateSystem;
        private bool _disposed;

        private LegacyXlsTabularWorkbook(LegacyBiffSource workbookStream, ExcelReadOptions options) {
            _workbookStream = workbookStream;
            ParseGlobals(
                workbookStream,
                options,
                out _sheets,
                out _sharedStrings,
                out _dateStyles,
                out _uses1904DateSystem);
            _tableNames = _sheets.Select(static sheet => sheet.Name).ToArray();
        }

        internal IReadOnlyList<string> TableNames => _tableNames;

        internal static LegacyXlsTabularWorkbook Open(
            string path,
            ExcelReadOptions options,
            CancellationToken cancellationToken = default) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (options == null) throw new ArgumentNullException(nameof(options));
            cancellationToken.ThrowIfCancellationRequested();
            var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                1,
                FileOptions.RandomAccess);
            return Open(stream, options, cancellationToken);
        }

        internal static LegacyXlsTabularWorkbook Open(
            byte[] bytes,
            ExcelReadOptions options,
            CancellationToken cancellationToken = default) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (options == null) throw new ArgumentNullException(nameof(options));
            cancellationToken.ThrowIfCancellationRequested();
            var stream = new MemoryStream(bytes, writable: false);
            return Open(stream, options, cancellationToken);
        }

        internal static IReadOnlyList<string> ReadSheetNames(
            string path,
            ExcelReadOptions options,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }
            if (options == null) throw new ArgumentNullException(nameof(options));

            cancellationToken.ThrowIfCancellationRequested();
            using var source = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                1,
                FileOptions.RandomAccess);
            if (source.Length > options.MaxInputBytes) {
                throw new InvalidDataException(
                    $"Workbook input contains {source.Length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
            }

            using Stream workbookStream = OpenWorkbookStream(source, options, cancellationToken);
            using var biffSource = new LegacyBiffSource(workbookStream, cancellationToken);
            IReadOnlyList<SheetInfo> sheets = ParseSheetNames(biffSource, options, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return sheets.Select(static sheet => sheet.Name).ToArray();
        }

        private static LegacyXlsTabularWorkbook Open(
            Stream source,
            ExcelReadOptions options,
            CancellationToken cancellationToken) {
            try {
                cancellationToken.ThrowIfCancellationRequested();
                if (source.Length > options.MaxInputBytes) {
                    throw new InvalidDataException(
                        $"Workbook input contains {source.Length} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
                }
                Stream workbookStream = OpenWorkbookStream(source, options, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                var biffSource = new LegacyBiffSource(workbookStream, cancellationToken);
                try {
                    return new LegacyXlsTabularWorkbook(biffSource, options);
                } catch {
                    biffSource.Dispose();
                    throw;
                }
            } catch {
                source.Dispose();
                throw;
            }
        }

        internal DbDataReader OpenTable(
            string tableName,
            bool hasHeaderRow,
            ExcelReadOptions options,
            CancellationToken cancellationToken = default) {
            ThrowIfDisposed();
            int sheetIndex = -1;
            for (int index = 0; index < _sheets.Count; index++) {
                if (string.Equals(_sheets[index].Name, tableName, StringComparison.OrdinalIgnoreCase)) {
                    sheetIndex = index;
                    break;
                }
            }
            if (sheetIndex < 0) {
                throw new KeyNotFoundException($"Worksheet '{tableName}' was not found.");
            }

            SheetInfo sheet = _sheets[sheetIndex];
            return new LegacyXlsTabularDataReader(
                _workbookStream,
                sheet.StreamOffset,
                _sharedStrings,
                _dateStyles,
                _uses1904DateSystem,
                hasHeaderRow,
                options,
                cancellationToken);
        }

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            _workbookStream.Dispose();
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(LegacyXlsTabularWorkbook));
        }

        private static Stream OpenWorkbookStream(
            Stream source,
            ExcelReadOptions options,
            CancellationToken cancellationToken) {
            if (!source.CanRead || !source.CanSeek) {
                throw new ArgumentException("The XLS fast path requires a readable seekable stream.", nameof(source));
            }

            long maximum = Math.Min(options.MaxInputBytes, int.MaxValue);
            var compoundOptions = new OfficeCompoundReadOptions(
                maxStreamBytes: maximum,
                maxTotalStreamBytes: maximum);
            bool read = OfficeCompoundFileReader.TryOpenStream(
                source,
                compoundOptions,
                static (name, _) => IsWorkbookStreamName(name),
                leaveOpen: false,
                cancellationToken,
                out Stream? workbookStream,
                out string? error);
            if (!read) {
                throw new InvalidDataException(error ?? "The input is not a supported OLE compound XLS file.");
            }
            if (workbookStream == null) {
                throw new InvalidDataException("The OLE compound file does not contain a Workbook or Book stream.");
            }
            return workbookStream;
        }

        private static bool IsWorkbookStreamName(string name) {
            return string.Equals(name, "Workbook", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Book", StringComparison.OrdinalIgnoreCase);
        }

        private static void ParseGlobals(
            LegacyBiffSource bytes,
            ExcelReadOptions options,
            out IReadOnlyList<SheetInfo> sheets,
            out IReadOnlyList<string> sharedStrings,
            out bool[] dateStyles,
            out bool uses1904DateSystem) {
            var parsedSheets = new List<SheetInfo>();
            var cellFormats = new List<XfInfo>();
            var customNumberFormats = new Dictionary<ushort, string>();
            var diagnostics = new List<LegacyXlsImportDiagnostic>();
            List<string>? parsedSharedStrings = null;
            int offset = 0;
            bool sawBof = false;
            bool sawEof = false;
            uses1904DateSystem = false;

            while (TryReadRecord(bytes, ref offset, out RecordSlice record)) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!sawBof) {
                    if (record.Type != (ushort)BiffRecordType.Bof || record.Length < 4) {
                        throw new InvalidDataException("The XLS workbook stream is missing a valid workbook-globals BOF record.");
                    }
                    ushort version = bytes.ReadUInt16(record.PayloadOffset);
                    ushort substreamType = bytes.ReadUInt16(record.PayloadOffset + 2);
                    if (version != Biff8Version || substreamType != 0x0005) {
                        throw new LegacyXlsFastPathNotSupportedException("The direct XLS reader supports BIFF8 workbook streams.");
                    }
                    sawBof = true;
                    continue;
                }

                switch ((BiffRecordType)record.Type) {
                    case BiffRecordType.BoundSheet8:
                        if (parsedSheets.Count >= options.MaxWorksheets) {
                            throw new InvalidDataException(
                                $"The XLS workbook contains more than the configured {options.MaxWorksheets} worksheet definitions.");
                        }
                        parsedSheets.Add(ReadBoundSheet(bytes, record));
                        break;
                    case BiffRecordType.Date1904:
                        if (record.Length < 2) throw Truncated(record, "Date1904");
                        uses1904DateSystem = bytes.ReadUInt16(record.PayloadOffset) != 0;
                        break;
                    case BiffRecordType.Format:
                        ReadNumberFormat(bytes, record, customNumberFormats);
                        break;
                    case BiffRecordType.Xf:
                        if (record.Length < 6) throw Truncated(record, "XF");
                        ushort protection = bytes.ReadUInt16(record.PayloadOffset + 4);
                        ushort attributes = record.Length >= 10
                            ? bytes.ReadUInt16(record.PayloadOffset + 8)
                            : (ushort)0;
                        cellFormats.Add(new XfInfo(
                            bytes.ReadUInt16(record.PayloadOffset + 2),
                            isStyle: (protection & 0x0004) != 0,
                            parentStyleIndex: (ushort)((protection >> 4) & 0x0fff),
                            applyNumberFormat: (attributes & 0x0400) != 0));
                        break;
                    case BiffRecordType.Sst:
                        if (parsedSharedStrings != null) {
                            throw new InvalidDataException("The XLS workbook contains more than one shared-string table.");
                        }
                        parsedSharedStrings = ReadSharedStrings(bytes, record, ref offset, options, diagnostics);
                        break;
                    case BiffRecordType.FilePass:
                        throw new LegacyXlsFastPathNotSupportedException("Encrypted XLS workbooks use the full legacy import path.");
                    case BiffRecordType.Eof:
                        sawEof = true;
                        offset = bytes.Length;
                        break;
                }
            }

            if (!sawBof || !sawEof) {
                throw new InvalidDataException("The XLS workbook-globals substream is truncated before EOF.");
            }

            SheetInfo[] worksheets = ValidateWorksheets(parsedSheets, bytes.Length);

            var styles = new bool[cellFormats.Count];
            for (int index = 0; index < styles.Length; index++) {
                ushort formatId = ResolveEffectiveNumberFormat(cellFormats, index);
                styles[index] = BiffBuiltInNumberFormat.IsDateLike(formatId)
                    || customNumberFormats.TryGetValue(formatId, out string? code)
                    && ExcelNumberFormatClassifier.LooksLikeDateFormat(code);
            }

            sheets = worksheets;
            sharedStrings = parsedSharedStrings != null
                ? parsedSharedStrings
                : Array.Empty<string>();
            dateStyles = styles;
        }

        private static IReadOnlyList<SheetInfo> ParseSheetNames(
            LegacyBiffSource bytes,
            ExcelReadOptions options,
            CancellationToken cancellationToken) {
            var boundSheetRecords = new List<RecordSlice>();
            int offset = 0;
            int sheetDefinitionCount = 0;
            ushort workbookVersion = 0;
            var workbookCodePage = new BiffCodePageState();
            bool sawBof = false;
            bool sawEof = false;
            while (TryReadRecord(bytes, ref offset, out RecordSlice record)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!sawBof) {
                    if (record.Type != (ushort)BiffRecordType.Bof || record.Length < 4) {
                        throw new InvalidDataException(
                            "The XLS workbook stream is missing a valid workbook-globals BOF record.");
                    }
                    workbookVersion = bytes.ReadUInt16(record.PayloadOffset);
                    ushort substreamType = bytes.ReadUInt16(record.PayloadOffset + 2);
                    if ((workbookVersion != Biff5Version && workbookVersion != Biff8Version)
                        || substreamType != 0x0005) {
                        throw new LegacyXlsFastPathNotSupportedException(
                            "Metadata-only sheet discovery supports BIFF5 and BIFF8 workbook streams.");
                    }
                    sawBof = true;
                    continue;
                }

                switch ((BiffRecordType)record.Type) {
                    case BiffRecordType.BoundSheet8:
                        if (sheetDefinitionCount >= options.MaxWorksheets) {
                            throw new InvalidDataException(
                                $"The XLS workbook contains more than the configured {options.MaxWorksheets} worksheet definitions.");
                        }
                        sheetDefinitionCount++;
                        boundSheetRecords.Add(record);
                        break;
                    case BiffRecordType.CodePage:
                        if (record.Length < 2) {
                            workbookCodePage.ObserveMalformed(record.Offset);
                        } else {
                            workbookCodePage.Observe(
                                bytes.ReadUInt16(record.PayloadOffset),
                                record.Offset);
                        }
                        break;
                    case BiffRecordType.FilePass:
                        throw new LegacyXlsFastPathNotSupportedException(
                            "Encrypted XLS workbooks are not supported by metadata-only sheet discovery.");
                    case BiffRecordType.Eof:
                        sawEof = true;
                        offset = bytes.Length;
                        break;
                }
            }

            if (!sawBof || !sawEof) {
                throw new InvalidDataException("The XLS workbook-globals substream is truncated before EOF.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            Encoding? sheetNameEncoding = workbookVersion == Biff5Version
                ? BiffCodePageEncoding.Resolve(workbookCodePage)
                : null;
            var parsedSheets = new List<SheetInfo>(boundSheetRecords.Count);
            foreach (RecordSlice record in boundSheetRecords) {
                cancellationToken.ThrowIfCancellationRequested();
                parsedSheets.Add(ReadBoundSheet(bytes, record, workbookVersion, sheetNameEncoding));
            }
            return ValidateWorksheets(parsedSheets, bytes.Length);
        }

        private static SheetInfo[] ValidateWorksheets(
            IReadOnlyList<SheetInfo> parsedSheets,
            int streamLength) {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            SheetInfo[] worksheets = parsedSheets
                .Where(static sheet => sheet.SheetType == 0)
                .ToArray();
            foreach (SheetInfo sheet in worksheets) {
                if (sheet.StreamOffset < 0 || sheet.StreamOffset >= streamLength) {
                    throw new InvalidDataException($"Worksheet '{sheet.Name}' has an invalid BIFF stream offset.");
                }
                if (!names.Add(sheet.Name)) {
                    throw new InvalidDataException(
                        $"The workbook contains duplicate worksheet name '{sheet.Name}' under case-insensitive matching.");
                }
            }
            if (worksheets.Length == 0) {
                throw new InvalidDataException("The workbook contains no readable worksheets.");
            }

            return worksheets;
        }

        private static ushort ResolveEffectiveNumberFormat(IReadOnlyList<XfInfo> cellFormats, int index) {
            XfInfo format = cellFormats[index];
            if (format.IsStyle
                || format.ApplyNumberFormat
                || format.ParentStyleIndex >= cellFormats.Count) {
                return format.NumberFormatId;
            }

            XfInfo parent = cellFormats[format.ParentStyleIndex];
            return parent.IsStyle ? parent.NumberFormatId : format.NumberFormatId;
        }

        private static List<string> ReadSharedStrings(
            LegacyBiffSource bytes,
            RecordSlice first,
            ref int offset,
            ExcelReadOptions options,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            var payloads = new List<byte[]> { CopyPayload(bytes, first) };
            int lookahead = offset;
            while (TryReadRecord(bytes, ref lookahead, out RecordSlice continuation)
                   && continuation.Type == (ushort)BiffRecordType.Continue) {
                payloads.Add(CopyPayload(bytes, continuation));
                offset = lookahead;
            }

            List<string> values = BiffStringReader.ReadSharedStringTexts(
                payloads,
                diagnostics,
                first.Offset,
                options.MaxSharedStringItems,
                options.MaxSharedStringItemCharacters,
                options.MaxSharedStringCharacters);
            if (diagnostics.Count > 0) {
                throw new InvalidDataException(diagnostics[diagnostics.Count - 1].Message);
            }
            if (values.Count > options.MaxSharedStringItems) {
                throw new InvalidDataException(
                    $"XLS shared-string count {values.Count} exceeds the configured limit of {options.MaxSharedStringItems}.");
            }
            long characters = 0;
            foreach (string value in values) {
                if (value.Length > options.MaxSharedStringItemCharacters) {
                    throw new InvalidDataException(
                        $"An XLS shared string contains {value.Length} characters, exceeding the configured limit of {options.MaxSharedStringItemCharacters}.");
                }
                characters = checked(characters + value.Length);
                if (characters > options.MaxSharedStringCharacters) {
                    throw new InvalidDataException(
                        $"XLS shared strings contain more than the configured {options.MaxSharedStringCharacters} characters.");
                }
            }
            return values;
        }

        private static SheetInfo ReadBoundSheet(
            LegacyBiffSource bytes,
            RecordSlice record,
            ushort workbookVersion = Biff8Version,
            Encoding? byteStringEncoding = null) {
            int minimumLength = workbookVersion == Biff5Version ? 7 : 8;
            if (record.Length < minimumLength) throw Truncated(record, "BoundSheet");
            byte[] payload = CopyPayload(bytes, record);
            int nameOffset = 6;
            string name = workbookVersion == Biff5Version
                ? BiffStringReader.ReadShortByteString(
                    payload,
                    ref nameOffset,
                    byteStringEncoding ?? BiffCodePageEncoding.Resolve((ushort?)null))
                : BiffStringReader.ReadShortUnicodeString(payload, ref nameOffset);
            if (string.IsNullOrWhiteSpace(name)) {
                throw new InvalidDataException("An XLS worksheet has an empty name.");
            }
            return new SheetInfo(
                name,
                checked((int)bytes.ReadUInt32(record.PayloadOffset)),
                bytes.ReadByte(record.PayloadOffset + 5));
        }

        private static void ReadNumberFormat(
            LegacyBiffSource bytes,
            RecordSlice record,
            IDictionary<ushort, string> formats) {
            if (record.Length < 5) throw Truncated(record, "Format");
            byte[] payload = CopyPayload(bytes, record);
            ushort id = BiffRecordReader.ReadUInt16(payload, 0);
            int valueOffset = 2;
            formats[id] = BiffStringReader.ReadUnicodeString(payload, ref valueOffset);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool TryReadRecord(LegacyBiffSource bytes, ref int offset, out RecordSlice record) {
            if (offset == bytes.Length) {
                record = default;
                return false;
            }
            if (offset < 0 || offset + 4 > bytes.Length) {
                throw new InvalidDataException("The BIFF stream ended inside a record header.");
            }
            ushort type = bytes.ReadUInt16(offset);
            ushort length = bytes.ReadUInt16(offset + 2);
            int payloadOffset = checked(offset + 4);
            int nextOffset = checked(payloadOffset + length);
            if (nextOffset > bytes.Length) {
                throw new InvalidDataException(
                    $"BIFF record 0x{type:X4} at offset {offset} declares {length} payload bytes, but the stream ends early.");
            }
            record = new RecordSlice(type, offset, payloadOffset, length);
            offset = nextOffset;
            return true;
        }

        internal static ushort ReadUInt16(LegacyBiffSource bytes, int offset) => bytes.ReadUInt16(offset);

        internal static uint ReadUInt32(LegacyBiffSource bytes, int offset) => bytes.ReadUInt32(offset);

        internal static double ReadDouble(LegacyBiffSource bytes, int offset) => bytes.ReadDouble(offset);

        internal static byte[] CopyPayload(LegacyBiffSource bytes, RecordSlice record) =>
            bytes.Copy(record.PayloadOffset, record.Length);

        private static InvalidDataException Truncated(RecordSlice record, string name) =>
            new($"The {name} record at offset {record.Offset} is truncated.");

        internal readonly struct RecordSlice {
            internal RecordSlice(ushort type, int offset, int payloadOffset, int length) {
                Type = type;
                Offset = offset;
                PayloadOffset = payloadOffset;
                Length = length;
            }

            internal ushort Type { get; }
            internal int Offset { get; }
            internal int PayloadOffset { get; }
            internal int Length { get; }
        }

        private readonly struct XfInfo {
            internal XfInfo(
                ushort numberFormatId,
                bool isStyle,
                ushort parentStyleIndex,
                bool applyNumberFormat) {
                NumberFormatId = numberFormatId;
                IsStyle = isStyle;
                ParentStyleIndex = parentStyleIndex;
                ApplyNumberFormat = applyNumberFormat;
            }

            internal ushort NumberFormatId { get; }

            internal bool IsStyle { get; }

            internal ushort ParentStyleIndex { get; }

            internal bool ApplyNumberFormat { get; }
        }

        private readonly struct SheetInfo {
            internal SheetInfo(string name, int streamOffset, byte sheetType) {
                Name = name;
                StreamOffset = streamOffset;
                SheetType = sheetType;
            }

            internal string Name { get; }
            internal int StreamOffset { get; }
            internal byte SheetType { get; }
        }
    }

    internal sealed class LegacyXlsFastPathNotSupportedException : NotSupportedException {
        internal LegacyXlsFastPathNotSupportedException(string message) : base(message) { }
    }
}
