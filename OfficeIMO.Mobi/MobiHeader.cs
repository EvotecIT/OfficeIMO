using System.Text;
using static OfficeIMO.Mobi.PalmDatabase;

namespace OfficeIMO.Mobi;

internal sealed class MobiHeader {
    private const int MaximumExpandedText = 16 * 1024 * 1024;
    private readonly Dictionary<int, List<byte[]>> _metadata = new();

    internal MobiHeader(PalmDatabase db, int record, CancellationToken token) {
        Record = record;
        ReadOnlySpan<byte> data = db.Record(record);
        Compression = U16(data, 0);
        if (!db.IsPalmDoc && U16(data, 12) != 0) RejectEncryption();
        Length = Number(data, 4);
        TextRecords = U16(data, 8);
        RecordSize = U16(data, 10);
        if (Length > MaximumExpandedText || TextRecords == 0 || TextRecords > db.Count - record - 1 ||
            RecordSize is 0 or > 4096 || Length > TextRecords * RecordSize)
            throw new InvalidDataException("The Mobipocket text size, record count or record size exceeds its bounds.");
        if (db.IsPalmDoc || data.Length == 16) {
            Version = db.IsPalmDoc ? 0 : 1;
            IsHeaderless = !db.IsPalmDoc;
            Encoding = GetEncoding(1252);
            return;
        }
        if (!HasSignature(Slice(data, 16, 4), "MOBI")) throw new InvalidDataException("The Mobipocket header signature is missing.");
        int size = Number(data, 20);
        if (size < 24) throw new InvalidDataException("The Mobipocket header is too short.");
        _ = Slice(data, 16, size);
        if (size >= 160 && U32(data, 168) is not (0 or uint.MaxValue) && U32(data, 172) is not (0 or uint.MaxValue))
            RejectEncryption();
        Version = Number(data, 36);
        if (Version is < 1 or > 8) throw new NotSupportedException("The Mobipocket book version is outside the supported range.");
        Encoding = GetEncoding(Number(data, 28));
        Locale = size >= 80 ? U32(data, 92) : 0;
        FirstImage = size >= 96 ? ReadReference(data, 108) : null;
        NavigationIndex = size >= 232 ? ReadReference(data, 244) : null;
        ExtraFlags = size >= 228 ? U16(data, 242) : 0;
        if (Version == 8) {
            FlowCount = size >= 184 ? Number(data, 196) : 0;
            if (FlowCount > 1) {
                FlowIndex = ReadReference(data, 192)
                    ?? throw new InvalidDataException("The Kindle book declares multiple flows without an FDST record.");
            }
            FragmentIndex = size >= 236 ? ReadReference(data, 248) : null;
            SkeletonIndex = size >= 240 ? ReadReference(data, 252) : null;
        }
        if (size >= 116 && (U32(data, 128) & 0x40) != 0) ReadExth(data.Slice(16 + size), token);
        if (size >= 76 && U32(data, 88) is > 0 and <= 16_384) {
            ReadOnlySpan<byte> bytes = Slice(data, Number(data, 84), Number(data, 88));
            HeaderTitle = Decode(Encoding, bytes);
        }
    }

    internal int Record { get; }
    internal int Compression { get; }
    internal int Length { get; }
    internal int TextRecords { get; }
    internal int RecordSize { get; }
    internal int Version { get; }
    internal bool IsKf8 => Version == 8;
    internal bool IsHeaderless { get; }
    internal Encoding Encoding { get; }
    internal uint Locale { get; }
    internal int ExtraFlags { get; }
    internal string? HeaderTitle { get; }
    internal int? FirstImage { get; }
    internal int? FlowIndex { get; }
    internal int FlowCount { get; }
    internal int? NavigationIndex { get; }
    internal int? FragmentIndex { get; }
    internal int? SkeletonIndex { get; }

    internal IEnumerable<string> Text(int id) => _metadata.TryGetValue(id, out List<byte[]>? values)
        ? values.Select(value => Decode(Encoding, value)).Where(value => value.Length != 0)
        : Enumerable.Empty<string>();

    internal int? Integer(int id) => _metadata.TryGetValue(id, out List<byte[]>? values) && values[0].Length == 4
        ? ReadReference(values[0], 0) : null;

    internal static Encoding GetEncoding(int codePage) {
        if (codePage == 65001) return new UTF8Encoding(false, true);
        if (codePage != 1252) throw new NotSupportedException("Mobipocket text must use UTF-8 or Windows-1252.");
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
    }

    internal static string Decode(Encoding encoding, ReadOnlySpan<byte> value) =>
        encoding.GetString(value.ToArray()).Trim('\0', ' ', '\r', '\n', '\t');

    private static int? ReadReference(ReadOnlySpan<byte> data, int offset) =>
        U32(data, offset) == uint.MaxValue ? null : Number(data, offset);

    private void ReadExth(ReadOnlySpan<byte> source, CancellationToken token) {
        if (!HasSignature(source, "EXTH")) throw new InvalidDataException("The declared Mobipocket EXTH header is missing.");
        int size = Number(source, 4);
        int count = Number(source, 8);
        if (size < 12 || size > 1024 * 1024 || count > 1_024)
            throw new InvalidDataException("The Mobipocket EXTH metadata exceeds its bounds.");
        source = Slice(source, 0, size);
        int cursor = 12;
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            int id = Number(source, cursor);
            int length = Number(source, cursor + 4);
            if (length < 8) throw new InvalidDataException("An EXTH item has an invalid length.");
            if (!_metadata.TryGetValue(id, out List<byte[]>? values)) _metadata.Add(id, values = new());
            values.Add(Slice(source, cursor + 8, length - 8).ToArray());
            cursor += length;
        }
    }

    private static void RejectEncryption() =>
        throw new NotSupportedException("DRM-protected Mobipocket and Kindle books are not supported.");
}
