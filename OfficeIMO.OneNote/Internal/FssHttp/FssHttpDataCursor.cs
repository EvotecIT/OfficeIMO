namespace OfficeIMO.OneNote;

internal sealed class FssHttpDataCursor {
    private readonly byte[] _data;
    private readonly long _absoluteOffset;
    private int _position;

    internal FssHttpDataCursor(byte[] data, long absoluteOffset) {
        _data = data;
        _absoluteOffset = absoluteOffset;
    }

    internal int Position => _position;
    internal int Remaining => _data.Length - _position;

    internal OneNoteExtendedGuid ReadExtendedGuid() {
        OneNoteExtendedGuid value;
        try { value = OneNoteExtendedGuidReader.Read(_data, _position); }
        catch (OneNoteFormatException exception) { throw Error(exception.Code, exception.Message); }
        _position += value.EncodedLength;
        return value;
    }

    internal void SkipSerialNumber() {
        Ensure(1);
        byte type = _data[_position];
        if (type == 0) { _position++; return; }
        if (type != 0x80) throw Error("ONENOTE_PACKAGE_SERIAL", "A serial number uses an invalid type prefix.");
        Ensure(25);
        _position += 25;
    }

    internal ulong ReadCompactUInt64() {
        Ensure(1);
        byte first = _data[_position++];
        if (first == 0) return 0;
        if (first == 0x80) {
            Ensure(8);
            ulong value = OneNoteBinary.ReadUInt64(_data, _position);
            _position += 8;
            return value;
        }
        int encodedLength = 1;
        byte mask = 1;
        while ((first & mask) == 0 && encodedLength < 8) { encodedLength++; mask <<= 1; }
        if (encodedLength > 7) throw Error("ONENOTE_PACKAGE_COMPACT_UINT", "A compact integer uses an invalid type prefix.");
        ulong raw = first;
        Ensure(encodedLength - 1);
        for (int index = 1; index < encodedLength; index++) raw |= (ulong)_data[_position++] << (index * 8);
        return raw >> encodedLength;
    }

    internal uint ReadUInt32() {
        Ensure(4);
        uint value = OneNoteBinary.ReadUInt32(_data, _position);
        _position += 4;
        return value;
    }

    internal Guid ReadGuid() {
        Ensure(16);
        Guid value = OneNoteBinary.ReadGuid(_data, _position);
        _position += 16;
        return value;
    }

    internal IReadOnlyList<OneNoteExtendedGuid> ReadExtendedGuidArray(int maxCount) {
        int count = CheckedCount(ReadCompactUInt64(), maxCount, "extended-GUID array");
        var values = new List<OneNoteExtendedGuid>(count);
        for (int index = 0; index < count; index++) values.Add(ReadExtendedGuid());
        return values.AsReadOnly();
    }

    internal IReadOnlyList<FssHttpCellId> ReadCellIdArray(int maxCount) {
        int count = CheckedCount(ReadCompactUInt64(), maxCount, "cell-ID array");
        var values = new List<FssHttpCellId>(count);
        for (int index = 0; index < count; index++) values.Add(ReadCellId());
        return values.AsReadOnly();
    }

    internal FssHttpCellId ReadCellId() => new FssHttpCellId(ReadExtendedGuid(), ReadExtendedGuid());

    internal byte[] ReadBinaryItem(long maxBytes) {
        ulong count = ReadCompactUInt64();
        if (count > (ulong)maxBytes || count > int.MaxValue) throw Error("ONENOTE_PACKAGE_BINARY_SIZE", "A binary item exceeds its configured materialization limit.");
        return ReadBytes((int)count);
    }

    internal byte[] ReadBytes(int count) {
        Ensure(count);
        var bytes = new byte[count];
        if (count > 0) Buffer.BlockCopy(_data, _position, bytes, 0, count);
        _position += count;
        return bytes;
    }

    internal void EnsureEnd(string name) {
        if (Remaining != 0) throw Error("ONENOTE_PACKAGE_ITEM_LENGTH", "The " + name + " contains unexpected trailing bytes.");
    }

    private int CheckedCount(ulong count, int maxCount, string name) {
        if (count > (ulong)maxCount || count > int.MaxValue) throw Error("ONENOTE_PACKAGE_ARRAY_LIMIT", "A " + name + " exceeds its configured count limit.");
        return (int)count;
    }

    private void Ensure(int count) {
        if (count < 0 || _position > _data.Length - count) throw Error("ONENOTE_PACKAGE_ITEM_BOUNDS", "A package item is truncated.");
    }

    private OneNoteFormatException Error(string code, string message) => new OneNoteFormatException(code, message, _absoluteOffset + _position);
}

internal readonly struct FssHttpCellId : IEquatable<FssHttpCellId> {
    internal FssHttpCellId(OneNoteExtendedGuid first, OneNoteExtendedGuid second) { First = first; Second = second; }
    internal OneNoteExtendedGuid First { get; }
    internal OneNoteExtendedGuid Second { get; }
    public bool Equals(FssHttpCellId other) => First.Equals(other.First) && Second.Equals(other.Second);
    public override bool Equals(object? obj) => obj is FssHttpCellId other && Equals(other);
    public override int GetHashCode() => (First.GetHashCode() * 397) ^ Second.GetHashCode();
}
