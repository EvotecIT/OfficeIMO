using OfficeIMO.Excel.LegacyXls.Biff;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Decodes bounded, pre-compound Excel worksheet and BIFF4 workbook streams.</summary>
internal sealed partial class EarlyBiffWorkbook {
    private const ushort Worksheet = 0x0010;
    private const ushort Workspace = 0x0100;
    private const ushort MacroSheet = 0x0040;
    private const ushort ChartSheet = 0x0020;
    private const ushort EndOfFile = 0x000A;
    private const ushort CodePage = 0x0042;
    private const ushort DateMode = 0x0022;
    private const ushort FilePass = 0x002F;
    private const ushort BoundSheet = 0x0085;
    private const ushort SheetsOffset = 0x008E;
    private const ushort SheetHeader = 0x008F;
    private const int HeaderSize = 4;
    private const int MaximumRows = 16_384;
    private const int MaximumColumns = 256;
    internal const int ProbeLength = 10;
    internal const int DefaultInputLimit = 64 * 1024 * 1024;

    internal static long? ProbeInputLimit(byte[] header, int length) =>
        Identify(header.AsSpan(0, length)) != null ? DefaultInputLimit : null;

    internal static long? ProbeBiff4WorkbookLimit(byte[] header, int length) =>
        Identify(header.AsSpan(0, length)) == "excel-biff4-workbook" ? DefaultInputLimit
            : throw new NotSupportedException("An XLW data reader requires an embedded BIFF4 workbook.");

    internal static Func<byte[], int, long?> ProbeInputLimitFor(int configuredLimit) => (header, length) =>
        Identify(header.AsSpan(0, length)) != null ? Math.Min(configuredLimit, DefaultInputLimit) : null;

    private readonly byte[] _data;
    private readonly OfficeLegacyImportLimits _limits;
    private readonly CancellationToken _cancellationToken;
    private readonly LegacySpreadsheetModel _model = new() {
        Quality = OfficeLegacyImportQuality.Structured,
        PreserveBlankCoordinates = true
    };
    private readonly HashSet<ushort> _unprojected = new();
    private int _records;
    private int _textCharacters;
    private bool _formulaOmitted;

    private EarlyBiffWorkbook(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        _data = data;
        _limits = limits;
        _cancellationToken = cancellationToken;
    }

    /// <summary>Returns a profile only for the expected native record header and worksheet or workspace kind.</summary>
    internal static string? Identify(ReadOnlySpan<byte> bytes) {
        if (bytes.Length < 8) return null;
        ushort record = BiffRecordReader.ReadUInt16(bytes, 0);
        ushort length = BiffRecordReader.ReadUInt16(bytes, 2);
        ushort kind = BiffRecordReader.ReadUInt16(bytes, 6);
        return (record, length, kind) switch {
            (0x0009, 4, Worksheet) => "excel-biff2-worksheet",
            (0x0209, 6, Worksheet) when bytes.Length >= ProbeLength => "excel-biff3-worksheet",
            (0x0409, 6, Worksheet) when bytes.Length >= ProbeLength => "excel-biff4-worksheet",
            (0x0409, 6, Workspace) when bytes.Length >= ProbeLength => "excel-biff4-workbook",
            _ => null
        };
    }

    internal static bool IsSource(string path, CancellationToken cancellationToken, bool requireBiff4Workbook = false) {
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        byte[] header = new byte[ProbeLength];
        int length = 0;
        while (length < header.Length) {
            int read = stream.Read(header, length, header.Length - length);
            if (read == 0) break;
            length += read;
        }
        string? profile = Identify(header.AsSpan(0, length));
        return requireBiff4Workbook ? profile == "excel-biff4-workbook" : profile != null;
    }

    internal static (LegacySpreadsheetModel Model, string Profile) Parse(
        byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length > limits.MaxInputBytes) {
            throw new InvalidDataException("Early BIFF input exceeds the configured byte limit.");
        }
        string profile = Identify(data) ?? throw new InvalidDataException("Unsupported early BIFF stream.");
        var parser = new EarlyBiffWorkbook(data, limits, cancellationToken);
        parser.Read();
        return (parser._model, profile);
    }

    private void Read() {
        int position = 0;
        Record bof = ReadRecord(ref position, _data.Length);
        if (ReadUInt16(bof, 2) == Workspace) {
            ReadWorkspace(position);
        } else {
            ReadWorksheet(bof.Type, position, _data.Length, "Sheet1", inheritedCodePage: null);
        }

        if (_model.Sheets.Count == 0) throw new InvalidDataException("The early BIFF workbook contains no supported worksheets.");
        if (_formulaOmitted) AddLoss("early-biff-formula-cache", "Formula", "Stored formula results were imported as values. Source expressions were omitted.");
        if (_unprojected.Count != 0) {
            string identifiers = string.Join(", ", _unprojected.OrderBy(static value => value).Take(24).Select(static value => $"0x{value:X4}"));
            if (_unprojected.Count > 24) identifiers += ", …";
            AddLoss("early-biff-records-omitted", "Structure", "Unprojected early BIFF record types: " + identifiers + ".");
        }
    }

    private void ReadWorkspace(int position) {
        _model.PreserveSheetNames = true;
        ushort? codePage = null;
        var catalog = new List<string>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int? firstHeader = null;
        while (position < _data.Length) {
            Record record = ReadRecord(ref position, _data.Length);
            switch (record.Type) {
                case EndOfFile:
                    RequireLength(record, 0);
                    if (position != _data.Length || firstHeader == null || names.Count != catalog.Count) {
                        throw new InvalidDataException("The BIFF4 workbook sheet catalog or trailing data does not match its embedded worksheets.");
                    }
                    return;
                case CodePage:
                    codePage = ObserveCodePage(record, codePage);
                    break;
                case SheetsOffset:
                    RequireLength(record, 4);
                    uint offset = BiffRecordReader.ReadUInt32(_data, record.Offset);
                    if (firstHeader != null || offset < position || offset >= _data.Length) {
                        throw new InvalidDataException("Invalid BIFF4 first-sheet offset.");
                    }
                    firstHeader = (int)offset;
                    break;
                case BoundSheet:
                    if (catalog.Count >= Math.Min(_limits.MaxItems, 256)) {
                        throw new InvalidDataException("The BIFF4 workbook exceeds its supported sheet count.");
                    }
                    catalog.Add(ReadText(record, 0, 1, BiffCodePageEncoding.Resolve(codePage)));
                    break;
                case SheetHeader:
                    if ((names.Count == 0 && firstHeader != record.Offset - HeaderSize) || names.Count >= catalog.Count) {
                        throw new InvalidDataException("The BIFF4 sheet header does not match the workbook catalog.");
                    }
                    string name = ReadText(record, 4, 1, BiffCodePageEncoding.Resolve(codePage));
                    if (!string.Equals(name, catalog[names.Count], StringComparison.Ordinal) || !names.Add(name)) {
                        throw new InvalidDataException("The BIFF4 workbook contains a mismatched or duplicate sheet name.");
                    }
                    uint length = BiffRecordReader.ReadUInt32(_data, record.Offset);
                    if (length > _data.Length - position) {
                        throw new InvalidDataException("The BIFF4 sheet extends beyond its workbook.");
                    }
                    int end = position + (int)length;
                    Record bof = ReadRecord(ref position, end);
                    RequireLength(bof, 6);
                    if (bof.Type != 0x0409) throw new InvalidDataException("A BIFF4 workbook sheet does not begin with a BIFF4 BOF.");
                    switch (ReadUInt16(bof, 2)) {
                        case Worksheet:
                            ReadWorksheet(bof.Type, position, end, name, codePage);
                            break;
                        case MacroSheet:
                            SkipNonWorksheet(position, end);
                            _model.InertContent |= OfficeLegacyInertContentKind.Macros;
                            AddLoss("early-biff-macro-sheet", "Active content", "An Excel 4 macro sheet was omitted without execution.");
                            break;
                        case ChartSheet:
                            SkipNonWorksheet(position, end);
                            AddLoss("early-biff-chart-sheet", "Chart", "An Excel 4 chart sheet was omitted.");
                            break;
                        default:
                            throw new InvalidDataException("The BIFF4 workbook contains an unknown embedded sheet kind.");
                    }
                    position = end;
                    break;
                default:
                    ObserveUnprojected(record);
                    break;
            }
        }
        throw new InvalidDataException("The BIFF4 workbook is missing its EOF record.");
    }

    private void SkipNonWorksheet(int position, int end) {
        while (position < end) {
            Record record = ReadRecord(ref position, end);
            if (IsBof(record.Type)) {
                SkipEmbeddedChart(record, ref position, end, charge: true);
                continue;
            }
            if (record.Type == FilePass) RejectEncrypted();
            if (record.Type != EndOfFile) continue;
            RequireLength(record, 0);
            if (position == end) return;
            break;
        }
        throw new InvalidDataException("An embedded BIFF4 sheet has an invalid EOF boundary.");
    }

    private static bool IsBof(ushort type) => type is 0x0009 or 0x0209 or 0x0409 or 0x0809;

    private void SkipEmbeddedChart(Record bof, ref int position, int end, bool charge) {
        if (ReadUInt16(bof, 2) != ChartSheet) throw new InvalidDataException("Unexpected nested BOF in an early BIFF worksheet.");
        if (charge) AddLoss("early-biff-embedded-chart", "Chart", "An embedded early BIFF chart was omitted.");
        while (position < end) {
            Record record = ReadRecord(ref position, end, charge);
            if (IsBof(record.Type)) throw new InvalidDataException("An embedded early BIFF chart contains an unexpected nested BOF.");
            if (record.Type == EndOfFile) {
                RequireLength(record, 0);
                return;
            }
            if (charge) ObserveUnprojected(record);
        }
        throw new InvalidDataException("An embedded early BIFF chart is missing its EOF record.");
    }

    private Record ReadRecord(ref int position, int end, bool charge = true) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (charge && ++_records > _limits.MaxRecords) {
            throw new InvalidDataException("Early BIFF input exceeds the configured record limit.");
        }
        if (position > end - HeaderSize) throw new InvalidDataException("Truncated early BIFF record header.");
        ushort type = BiffRecordReader.ReadUInt16(_data, position);
        int length = BiffRecordReader.ReadUInt16(_data, position + 2);
        int payload = position + HeaderSize;
        if (length > end - payload) throw new InvalidDataException($"Truncated early BIFF record 0x{type:X4}.");
        position = payload + length;
        return new Record(type, payload, length);
    }

    private ushort? ObserveCodePage(Record record, ushort? previous) {
        RequireLength(record, 2);
        ushort value = ReadUInt16(record, 0);
        if (previous.HasValue && previous.Value != value) throw new InvalidDataException("Conflicting early BIFF code page declarations.");
        return value;
    }

    private string ReadText(Record record, int offset, int lengthSize, Encoding encoding) {
        if (record.Length < offset + lengthSize) throw new InvalidDataException("Truncated early BIFF text length.");
        int count = lengthSize == 1 ? _data[record.Offset + offset] : ReadUInt16(record, offset);
        if (record.Length != offset + lengthSize + count) throw new InvalidDataException("Early BIFF text length does not match its record.");
        try {
            string value = encoding.GetString(_data, record.Offset + offset + lengthSize, count);
            if (value.Length > _limits.MaxTextCharacters - _textCharacters) {
                throw new InvalidDataException("Early BIFF input exceeds the configured text limit.");
            }
            _textCharacters += value.Length;
            return value;
        } catch (DecoderFallbackException error) {
            throw new InvalidDataException("An early BIFF text record is not valid in the declared code page.", error);
        }
    }

    private void ObserveUnprojected(Record record) {
        if (record.Type == FilePass) RejectEncrypted();
        bool ownSheet = record.Type == 0x0017 && record.Length >= 3
            && _data[record.Offset + 1] == 0x03 && _data[record.Offset] + 2 == record.Length;
        if (!ownSheet && record.Type is 0x0017 or 0x0023 or 0x0223) {
            _model.InertContent |= OfficeLegacyInertContentKind.ExternalLinks;
        }
        if (record.Type == 0x005D) _model.InertContent |= OfficeLegacyInertContentKind.EmbeddedObjects;
        _unprojected.Add(record.Type);
    }

    private void AddLoss(string code, string category, string message) {
        if (_model.Findings.Any(finding => finding.Code == code)) return;
        _model.Findings.Add(LegacySpreadsheetAdapterBase.LossFinding(code, category, message));
    }

    private ushort ReadUInt16(Record record, int offset) {
        if (offset > record.Length - 2) throw new InvalidDataException($"Truncated early BIFF record 0x{record.Type:X4}.");
        return BiffRecordReader.ReadUInt16(_data, record.Offset + offset);
    }

    private static void RequireLength(Record record, int expected) {
        if (record.Length != expected) throw new InvalidDataException($"Invalid early BIFF record length for 0x{record.Type:X4}.");
    }

    private static void RejectEncrypted() =>
        throw new NotSupportedException("Encrypted early BIFF worksheets are outside the supported import profile.");

    private readonly struct Record {
        internal Record(ushort type, int offset, int length) {
            Type = type;
            Offset = offset;
            Length = length;
        }
        internal ushort Type { get; }
        internal int Offset { get; }
        internal int Length { get; }
    }
}
