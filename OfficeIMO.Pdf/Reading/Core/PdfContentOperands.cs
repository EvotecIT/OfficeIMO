using System.Collections;

namespace OfficeIMO.Pdf;

/// <summary>
/// Typed operand list for synchronous content visitors. Numbers stay unboxed and strings stay in
/// a reusable byte buffer until a caller reads them as objects; the object view then returns the
/// same boxed number or byte array that an eager list would have held, and keeps returning it.
/// </summary>
internal sealed class PdfContentOperands : IReadOnlyList<object> {
    private const byte ObjectKind = 0;
    private const byte NumberKind = 1;
    private const byte StringKind = 2;
    private const int MaximumRetainedStringBytes = 64 * 1024;

    private object?[] _values;
    private double[] _numbers;
    private int[] _stringOffsets;
    private int[] _stringLengths;
    private byte[] _kinds;
    private byte[] _strings = Array.Empty<byte>();
    private int _stringsLength;
    private int _count;

    internal PdfContentOperands(int capacity = 8) {
        capacity = Math.Max(1, capacity);
        _values = new object?[capacity];
        _numbers = new double[capacity];
        _stringOffsets = new int[capacity];
        _stringLengths = new int[capacity];
        _kinds = new byte[capacity];
    }

    public int Count => _count;

    /// <summary>Operand slots currently allocated, for callers that cap retained buffers.</summary>
    internal int Capacity => _values.Length;

    /// <summary>Object view matching an eager operand list: numbers as <c>double</c>, strings as <c>byte[]</c>, names and arrays.</summary>
    public object this[int index] {
        get {
            if ((uint)index >= (uint)_count) ThrowIndexOutOfRange(index);
            object? value = _values[index];
            if (value is not null) return value;
            value = _kinds[index] == NumberKind
                ? PdfContentNumberBoxes.Box(_numbers[index])
                : CopyString(index);
            _values[index] = value;
            return value;
        }
    }

    public IEnumerator<object> GetEnumerator() {
        for (int index = 0; index < _count; index++) yield return this[index];
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    internal void Clear() {
        Array.Clear(_values, 0, _count);
        _count = 0;
        _stringsLength = 0;
        // A long inline string should not pin its buffer for the life of a reused list.
        if (_strings.Length > MaximumRetainedStringBytes) _strings = Array.Empty<byte>();
    }

    internal void AddObject(object value) {
        int index = Reserve();
        _values[index] = value;
        _kinds[index] = ObjectKind;
    }

    internal void AddNumber(double value) {
        int index = Reserve();
        _numbers[index] = value;
        _kinds[index] = NumberKind;
    }

    /// <summary>Reserves string storage; the caller writes at most <paramref name="maximumLength"/> bytes, then commits.</summary>
    internal byte[] BeginString(int maximumLength, out int offset) {
        offset = _stringsLength;
        int required = checked(_stringsLength + maximumLength);
        if (required > _strings.Length) {
            byte[] grown = new byte[Math.Max(required, Math.Max(64, _strings.Length * 2))];
            Buffer.BlockCopy(_strings, 0, grown, 0, _stringsLength);
            _strings = grown;
        }
        return _strings;
    }

    internal void CommitString(int offset, int length) {
        int index = Reserve();
        _stringOffsets[index] = offset;
        _stringLengths[index] = length;
        _kinds[index] = StringKind;
        _stringsLength = offset + length;
    }

    internal void RemoveLast() {
        _count--;
        _values[_count] = null;
        if (_kinds[_count] == StringKind) _stringsLength = _stringOffsets[_count];
    }

    /// <summary>True for an operand an eager list would hold as a boxed <see cref="double"/>.</summary>
    internal bool TryGetNumber(int index, out double value) {
        if ((uint)index >= (uint)_count) ThrowIndexOutOfRange(index);
        if (_kinds[index] == NumberKind) {
            value = _numbers[index];
            return true;
        }
        if (_values[index] is double number) {
            value = number;
            return true;
        }
        value = 0D;
        return false;
    }

    /// <summary>The operand as a number, or zero when it is not one.</summary>
    internal double NumberOrZero(int index) => TryGetNumber(index, out double value) ? value : 0D;

    /// <summary>True for an operand an eager list would hold as a <c>byte[]</c>.</summary>
    internal bool IsString(int index) {
        if ((uint)index >= (uint)_count) ThrowIndexOutOfRange(index);
        return _kinds[index] == StringKind && _values[index] is null || _values[index] is byte[];
    }

    internal int StringLength(int index) {
        if ((uint)index >= (uint)_count) ThrowIndexOutOfRange(index);
        return _values[index] is byte[] bytes ? bytes.Length : _stringLengths[index];
    }

    internal void CopyStringTo(int index, byte[] destination, int destinationOffset) {
        if ((uint)index >= (uint)_count) ThrowIndexOutOfRange(index);
        if (_values[index] is byte[] bytes) {
            Buffer.BlockCopy(bytes, 0, destination, destinationOffset, bytes.Length);
        } else {
            Buffer.BlockCopy(_strings, _stringOffsets[index], destination, destinationOffset, _stringLengths[index]);
        }
    }

    internal bool AllNumbers() {
        for (int index = 0; index < _count; index++) {
            if (!TryGetNumber(index, out _)) return false;
        }
        return true;
    }

    internal double[] ToNumberArray() {
        var numbers = new double[_count];
        for (int index = 0; index < numbers.Length; index++) TryGetNumber(index, out numbers[index]);
        return numbers;
    }

    internal object[] ToObjectArray() {
        if (_count == 0) return Array.Empty<object>();
        var values = new object[_count];
        for (int index = 0; index < values.Length; index++) values[index] = this[index];
        return values;
    }

    internal List<object> ToList() {
        var values = new List<object>(_count);
        for (int index = 0; index < _count; index++) values.Add(this[index]);
        return values;
    }

    /// <summary>Replaces this list with the operands of <paramref name="source"/>, keeping them typed.</summary>
    internal void CopyFrom(PdfContentOperands source) {
        Clear();
        EnsureCapacity(source._count);
        if (source._stringsLength > _strings.Length) _strings = new byte[Math.Max(source._stringsLength, _strings.Length * 2)];
        Buffer.BlockCopy(source._strings, 0, _strings, 0, source._stringsLength);
        _stringsLength = source._stringsLength;
        Array.Copy(source._values, _values, source._count);
        Array.Copy(source._numbers, _numbers, source._count);
        Array.Copy(source._stringOffsets, _stringOffsets, source._count);
        Array.Copy(source._stringLengths, _stringLengths, source._count);
        Array.Copy(source._kinds, _kinds, source._count);
        _count = source._count;
    }

    /// <summary>Appends operands from another visitor's list, keeping typed operands unboxed.</summary>
    internal void AddRange(IReadOnlyList<object> source) {
        if (source is PdfContentOperands typed) {
            for (int index = 0; index < typed._count; index++) {
                if (typed._values[index] is object value) {
                    AddObject(value);
                } else if (typed._kinds[index] == NumberKind) {
                    AddNumber(typed._numbers[index]);
                } else {
                    int length = typed._stringLengths[index];
                    byte[] buffer = BeginString(length, out int offset);
                    Buffer.BlockCopy(typed._strings, typed._stringOffsets[index], buffer, offset, length);
                    CommitString(offset, length);
                }
            }
            return;
        }
        for (int index = 0; index < source.Count; index++) AddObject(source[index]);
    }

    private byte[] CopyString(int index) {
        int length = _stringLengths[index];
        if (length == 0) return Array.Empty<byte>();
        var bytes = new byte[length];
        Buffer.BlockCopy(_strings, _stringOffsets[index], bytes, 0, length);
        return bytes;
    }

    private int Reserve() {
        if (_count == _values.Length) EnsureCapacity(_count + 1);
        return _count++;
    }

    private void EnsureCapacity(int capacity) {
        if (capacity <= _values.Length) return;
        int size = Math.Max(capacity, _values.Length * 2);
        Array.Resize(ref _values, size);
        Array.Resize(ref _numbers, size);
        Array.Resize(ref _stringOffsets, size);
        Array.Resize(ref _stringLengths, size);
        Array.Resize(ref _kinds, size);
    }

    private static void ThrowIndexOutOfRange(int index) =>
        throw new ArgumentOutOfRangeException(nameof(index), "Index was out of range. Must be non-negative and less than the size of the collection.");
}

/// <summary>Shares boxes for repeated operand numbers; boxed doubles are immutable.</summary>
internal static class PdfContentNumberBoxes {
    // Coordinates and kerning values recur across the content streams of a document, so the cache outlives one reader.
    [ThreadStatic]
    private static object?[]? t_boxedNumbers;
    private const int BoxedNumberSlots = 1024;

    internal static object Box(double value) {
        long bits = BitConverter.DoubleToInt64Bits(value);
        int slot = (int)((ulong)(bits ^ (bits >> 29) ^ (bits >> 47)) % BoxedNumberSlots);
        object?[] boxedNumbers = t_boxedNumbers ??= new object?[BoxedNumberSlots];
        if (boxedNumbers[slot] is double cached && BitConverter.DoubleToInt64Bits(cached) == bits) {
            return boxedNumbers[slot]!;
        }
        object boxed = value;
        boxedNumbers[slot] = boxed;
        return boxed;
    }
}
