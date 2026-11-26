using System.Text;
using static OfficeIMO.Mobi.PalmDatabase;

namespace OfficeIMO.Mobi;

// One INDX header owns TAGX, the IDXT-addressed data records and their following string records.
internal sealed class MobiIndex {
    private const int MaximumEntries = 12_000;
    private const int MaximumIndexRecords = 128;
    private readonly PalmDatabase _database;
    private readonly Encoding _encoding;
    private readonly int _firstStringRecord;
    private readonly int _stringRecords;

    internal MobiIndex(PalmDatabase database, int record, Encoding encoding, CancellationToken token) {
        _database = database;
        _encoding = encoding;
        ReadOnlySpan<byte> head = database.Record(record);
        if (!HasSignature(head, "INDX")) throw new InvalidDataException("The Mobipocket INDX record is missing.");
        int records = Number(head, 24);
        int expected = Number(head, 36);
        int stringRecords = Number(head, 52);
        if (records > MaximumIndexRecords || stringRecords > MaximumIndexRecords || expected > MaximumEntries ||
            1 + records + stringRecords > database.Count - record)
            throw new InvalidDataException("The Mobipocket INDX record counts exceed their bounds.");
        _firstStringRecord = record + records + 1;
        _stringRecords = stringRecords;
        int tagStart = Number(head, 4);
        ReadOnlySpan<byte> tag = Slice(head, tagStart, head.Length - tagStart);
        if (!HasSignature(tag, "TAGX")) throw new InvalidDataException("The Mobipocket TAGX control table is missing.");
        int tagLength = Number(tag, 4);
        int controls = Number(tag, 8);
        if (tagLength is < 16 or > 512 || tagLength % 4 != 0 || controls is < 1 or > 32)
            throw new InvalidDataException("The Mobipocket TAGX control table is invalid.");
        ReadOnlySpan<byte> tags = Slice(tag, 12, tagLength - 12);

        var rows = new List<Row>(expected);
        for (int i = 1; i <= records; i++) {
            token.ThrowIfCancellationRequested();
            ReadOnlySpan<byte> data = database.Record(record + i);
            if (!HasSignature(data, "INDX")) throw new InvalidDataException("A Mobipocket INDX data record is missing.");
            int directory = Number(data, 20);
            int number = Number(data, 24);
            if (number > expected - rows.Count || !HasSignature(Slice(data, directory, 4), "IDXT"))
                throw new InvalidDataException("A Mobipocket IDXT record directory is invalid.");
            ReadOnlySpan<byte> positions = Slice(data, directory + 4, number * 2);
            int previous = Number(data, 4);
            for (int j = 0; j < number; j++) {
                token.ThrowIfCancellationRequested();
                int start = U16(positions, j * 2);
                int end = j + 1 == number ? directory : U16(positions, j * 2 + 2);
                if (start < previous || start >= end || end > directory)
                    throw new InvalidDataException("A Mobipocket index entry is outside its IDXT boundaries.");
                rows.Add(ReadRow(data.Slice(start, end - start), tags, controls));
                previous = end;
            }
        }
        if (rows.Count != expected) throw new InvalidDataException("The Mobipocket INDX entry count does not match its header.");
        Rows = rows;
    }

    internal IReadOnlyList<Row> Rows { get; }

    internal string ReadLabel(int address) {
        int record = address >> 16;
        int offset = address & 0xffff;
        if (record >= _stringRecords) throw new InvalidDataException("A Mobipocket index label points outside its string records.");
        ReadOnlySpan<byte> data = _database.Record(_firstStringRecord + record);
        int length = Vwi(data, ref offset);
        if (length > 4_096) throw new InvalidDataException("A Mobipocket index label exceeds 4,096 bytes.");
        return MobiHeader.Decode(_encoding, Slice(data, offset, length));
    }

    private Row ReadRow(ReadOnlySpan<byte> data, ReadOnlySpan<byte> tags, int controls) {
        int nameLength = data[0];
        string name = MobiHeader.Decode(_encoding, Slice(data, 1, nameLength));
        ReadOnlySpan<byte> control = Slice(data, 1 + nameLength, controls);
        int position = 1 + nameLength + controls;
        int controlIndex = 0;
        var active = new List<(int Tag, int Count, int ByteLength)>();
        for (int i = 0; i < tags.Length; i += 4) {
            if (tags[i + 3] == 1) {
                controlIndex++;
                continue;
            }
            if (controlIndex >= control.Length || tags[i + 1] == 0 || tags[i + 2] == 0)
                throw new InvalidDataException("A Mobipocket TAGX definition is invalid.");
            int mask = tags[i + 2];
            int value = control[controlIndex] & mask;
            if (value == 0) continue;
            if (value == mask && (mask & (mask - 1)) != 0) {
                active.Add((tags[i], 0, Vwi(data, ref position)));
                continue;
            }
            while ((mask & 1) == 0) { mask >>= 1; value >>= 1; }
            active.Add((tags[i], value * tags[i + 1], 0));
        }
        var values = new Dictionary<int, int[]>();
        foreach (var definition in active) {
            if (definition.Count > 128 || definition.ByteLength > 512)
                throw new InvalidDataException("A Mobipocket index tag exceeds its value bounds.");
            int end = position + definition.ByteLength;
            if (end > data.Length) throw new InvalidDataException("A Mobipocket index value exceeds its entry.");
            var decoded = new List<int>();
            if (definition.ByteLength > 0) {
                while (position < end) decoded.Add(Vwi(data.Slice(0, end), ref position));
            } else {
                for (int i = 0; i < definition.Count; i++) decoded.Add(Vwi(data, ref position));
            }
            if (!values.ContainsKey(definition.Tag)) values.Add(definition.Tag, decoded.ToArray());
        }
        return new Row(name, values);
    }

    private static int Vwi(ReadOnlySpan<byte> data, ref int offset) {
        long number = 0;
        for (int i = 0; i < 5 && offset < data.Length; i++) {
            byte value = data[offset++];
            number = (number << 7) | (uint)(value & 0x7f);
            if ((value & 0x80) != 0 && number <= int.MaxValue) return (int)number;
        }
        throw new InvalidDataException("A Mobipocket variable-width integer is incomplete or out of range.");
    }

    internal sealed record Row(string Name, Dictionary<int, int[]> Tags) {
        internal int? Get(int tag, int index = 0) => Tags.TryGetValue(tag, out int[]? values) && index < values.Length ? values[index] : null;
        internal int Require(int tag, int index = 0) => Get(tag, index) ??
            throw new InvalidDataException("A required Kindle reconstruction field is absent from the INDX entry.");
    }
}
