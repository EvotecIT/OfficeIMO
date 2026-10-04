namespace OfficeIMO.Email.Store;

/// <summary>Reads the bplist00 object and offset tables without native libraries.</summary>
internal sealed class EmlxBinaryPlistReader {
    private readonly byte[] _data;
    private readonly int _start;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly int[] _offsets;
    private readonly int _referenceWidth;
    private readonly int _objectEnd;
    private readonly int _root;
    private readonly HashSet<int> _active = new HashSet<int>();
    private readonly Dictionary<int, object?> _scalars = new Dictionary<int, object?>();
    private int _propertyCount;

    internal EmlxBinaryPlistReader(byte[] data, int start, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _data = data;
        _start = start;
        _options = options;
        _cancellationToken = cancellationToken;
        if (data.Length - start < 40) throw Invalid("The binary plist is shorter than its header and trailer.");
        int trailer = data.Length - 32;
        int offsetWidth = data[trailer + 6];
        _referenceWidth = data[trailer + 7];
        if (offsetWidth < 1 || offsetWidth > 8 || _referenceWidth < 1 || _referenceWidth > 8) {
            throw Invalid("The binary plist uses an invalid offset or reference width.");
        }
        ulong count = Unsigned(trailer + 8, 8, data.Length);
        ulong root = Unsigned(trailer + 16, 8, data.Length);
        ulong table = Unsigned(trailer + 24, 8, data.Length);
        long maximumObjects = 2L * options.MaxPropertiesPerItem + 1;
        if (count > (ulong)maximumObjects) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                count > long.MaxValue ? long.MaxValue : (long)count, maximumObjects);
        }
        if (count == 0 || root >= count || table < 8 || table > (ulong)(trailer - start) ||
            count > (ulong)(trailer - start - (long)table) / (uint)offsetWidth) {
            throw Invalid("The binary plist offset table or root reference is outside the source.");
        }
        _objectEnd = start + (int)table;
        _root = (int)root;
        _offsets = new int[(int)count];
        for (int index = 0; index < _offsets.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ulong value = Unsigned(_objectEnd + index * offsetWidth, offsetWidth, trailer);
            if (value < 8 || value >= table) throw Invalid("A binary plist object offset is outside the object table.");
            _offsets[index] = start + (int)value;
        }
    }

    internal IReadOnlyDictionary<string, object?> Read() =>
        Parse(_root, 0) as Dictionary<string, object?> ?? throw Invalid("The binary plist root is not a dictionary.");

    private object? Parse(int reference, int depth) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > _options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth), depth, _options.MaxBTreeDepth);
        }
        if (reference < 0 || reference >= _offsets.Length) throw Invalid("A binary plist reference is outside the object table.");
        if (_scalars.TryGetValue(reference, out object? cached)) return cached;
        if (!_active.Add(reference)) throw Invalid("The binary plist contains a cyclic object graph.");
        try {
            int position = _offsets[reference];
            byte marker = _data[position++];
            int type = marker >> 4;
            int info = marker & 15;
            object? value;
            if (marker == 0x08 || marker == 0x09) value = marker == 0x09;
            else if (type == 1) {
                if (info > 3) throw new NotSupportedException("Binary plist integers wider than Int64 are retained as opaque metadata.");
                value = unchecked((long)Unsigned(position, 1 << info, _objectEnd));
            } else if (type == 2 || marker == 0x33) {
                int width = marker == 0x33 ? 8 : info == 2 ? 4 : info == 3 ? 8 : 0;
                if (width == 0) throw Invalid("A binary plist real has an invalid width.");
                ulong bits = Unsigned(position, width, _objectEnd);
                double number = width == 8 ? BitConverter.Int64BitsToDouble(unchecked((long)bits)) :
                    BitConverter.ToSingle(BitConverter.GetBytes(unchecked((uint)bits)), 0);
                if (double.IsNaN(number) || double.IsInfinity(number)) throw Invalid("A binary plist number is not finite.");
                if (marker == 0x33) {
                    try { value = new DateTimeOffset(2001, 1, 1, 0, 0, 0, TimeSpan.Zero).AddSeconds(number); }
                    catch (ArgumentOutOfRangeException) { throw Invalid("A binary plist date is outside the supported range."); }
                } else value = number;
            } else if (type == 4 || type == 5 || type == 6) {
                int length = Length(info, ref position);
                long bytes = type == 6 ? 2L * length : length;
                Require(position, bytes, _objectEnd);
                if (type == 4) {
                    var content = new byte[length];
                    Buffer.BlockCopy(_data, position, content, 0, length);
                    value = content;
                } else {
                    try {
                        value = (type == 6 ? new UnicodeEncoding(true, false, true) :
                            Encoding.GetEncoding("us-ascii", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback))
                            .GetString(_data, position, (int)bytes);
                    } catch (DecoderFallbackException) { throw Invalid("A binary plist string has invalid encoding."); }
                }
            } else if (type == 10 || type == 13) {
                int length = Length(info, ref position);
                Require(position, (long)length * _referenceWidth * (type == 13 ? 2 : 1), _objectEnd);
                if (length > _options.MaxPropertiesPerItem - _propertyCount) {
                    throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                        _propertyCount + (long)length, _options.MaxPropertiesPerItem);
                }
                _propertyCount += length;
                if (type == 10) {
                    var array = new object?[length];
                    for (int index = 0; index < length; index++) array[index] = Parse(Reference(position + index * _referenceWidth), depth + 1);
                    return array;
                }
                var dictionary = new Dictionary<string, object?>(StringComparer.Ordinal);
                for (int index = 0; index < length; index++) {
                    string key = Parse(Reference(position + index * _referenceWidth), depth + 1) as string ??
                        throw Invalid("A binary plist dictionary key is not a string.");
                    if (dictionary.ContainsKey(key)) throw Invalid("A binary plist dictionary contains a duplicate key.");
                    dictionary.Add(key, Parse(Reference(position + (length + index) * _referenceWidth), depth + 1));
                }
                return dictionary;
            } else {
                throw new NotSupportedException("The binary plist contains an object type that is retained as opaque metadata.");
            }
            _scalars.Add(reference, value);
            return value;
        } finally { _active.Remove(reference); }
    }

    private int Reference(int position) {
        ulong value = Unsigned(position, _referenceWidth, _objectEnd);
        if (value >= (ulong)_offsets.Length) throw Invalid("A binary plist reference is outside the object table.");
        return (int)value;
    }

    private int Length(int info, ref int position) {
        if (info < 15) return info;
        Require(position, 1, _objectEnd);
        byte marker = _data[position++];
        if ((marker >> 4) != 1 || (marker & 15) > 3) throw Invalid("A binary plist object length is not a bounded integer.");
        int width = 1 << (marker & 15);
        ulong value = Unsigned(position, width, _objectEnd);
        position += width;
        if (value > int.MaxValue) throw Invalid("A binary plist object length exceeds the addressable source.");
        return (int)value;
    }

    private ulong Unsigned(int position, int width, int end) {
        Require(position, width, end);
        ulong value = 0;
        for (int index = 0; index < width; index++) value = (value << 8) | _data[position + index];
        return value;
    }

    private void Require(int position, long count, int end) {
        if (position < _start || count < 0 || count > end - (long)position) {
            throw Invalid("A binary plist object extends outside its source table.");
        }
    }

    private static InvalidDataException Invalid(string message) => new InvalidDataException(message);
}
