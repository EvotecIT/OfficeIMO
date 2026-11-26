using System.Buffers.Binary;

namespace OfficeIMO.Mobi;

internal sealed class PalmDatabase {
    private const int HeaderLength = 78;
    private const int MaximumRecords = 16_384;
    private readonly byte[] _data;
    private readonly int[] _offsets;

    internal PalmDatabase(byte[] data, CancellationToken token) {
        _data = data;
        if (data.Length < HeaderLength || U32(data, 72) != 0)
            throw new InvalidDataException("The Palm database header or record directory is invalid.");
        IsPalmDoc = HasSignature(data.AsSpan(60), "TEXtREAd");
        if (!IsPalmDoc && !HasSignature(data.AsSpan(60), "BOOKMOBI"))
            throw new NotSupportedException("The Palm database is neither Mobipocket nor PalmDOC.");
        int count = U16(data, 76);
        if (count is < 2 or > MaximumRecords || HeaderLength + count * 8 > data.Length)
            throw new InvalidDataException("The Palm database record count or directory length exceeds its bounds.");
        int minimum = HeaderLength + count * 8;
        _offsets = new int[count + 1];
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            int position = Number(data, HeaderLength + i * 8);
            if (position < minimum || position > data.Length)
                throw new InvalidDataException("The Palm database record offsets are not ordered within the file.");
            _offsets[i] = position;
            minimum = position;
        }
        _offsets[count] = data.Length;
    }

    internal bool IsPalmDoc { get; }
    internal int Count => _offsets.Length - 1;
    internal ReadOnlySpan<byte> Name => _data.AsSpan(0, 32);

    internal ReadOnlySpan<byte> Record(int index) {
        if (index < 0 || index >= Count) throw new InvalidDataException("A Palm database record index is out of range.");
        return _data.AsSpan(_offsets[index], _offsets[index + 1] - _offsets[index]);
    }

    internal static bool HasSignature(ReadOnlySpan<byte> data, string signature) {
        if (data.Length < signature.Length) return false;
        for (int i = 0; i < signature.Length; i++)
            if (data[i] != signature[i]) return false;
        return true;
    }

    internal static ushort U16(ReadOnlySpan<byte> data, int offset) => BinaryPrimitives.ReadUInt16BigEndian(Slice(data, offset, 2));
    internal static uint U32(ReadOnlySpan<byte> data, int offset) => BinaryPrimitives.ReadUInt32BigEndian(Slice(data, offset, 4));

    internal static int Number(ReadOnlySpan<byte> data, int offset) {
        uint value = U32(data, offset);
        if (value > int.MaxValue) throw new InvalidDataException("A Mobipocket number exceeds its supported range.");
        return (int)value;
    }

    internal static ReadOnlySpan<byte> Slice(ReadOnlySpan<byte> data, int offset, int length) {
        if (offset < 0 || length < 0 || offset > data.Length - length)
            throw new InvalidDataException("A Mobipocket field extends beyond its record.");
        return data.Slice(offset, length);
    }
}
