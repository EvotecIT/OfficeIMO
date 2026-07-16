using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstPropertyContextReader {
    private readonly PstHeap _heap;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;

    static PstPropertyContextReader() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal PstPropertyContextReader(PstHeap heap, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _heap = heap;
        _options = options;
        _cancellationToken = cancellationToken;
    }

    internal List<MapiProperty> ReadProperties(IDictionary<ushort, uint>? sourceHnids = null,
        ISet<ushort>? includedPropertyIds = null,
        ISet<ushort>? deferredPropertyIds = null,
        long? maximumDecodedBytes = null) {
        if (_heap.ClientSignature != 0xBC) {
            throw new InvalidDataException("The PST node is not a Property Context.");
        }
        byte[] header = _heap.GetAllocation(_heap.UserRoot);
        if (header.Length < 8 || header[0] != 0xB5 || header[1] != 2 || header[2] != 6) {
            throw new InvalidDataException("The PST Property Context BTH header is invalid.");
        }

        int indexLevels = header[3];
        uint rootHid = PstBinary.UInt32(header, 4);
        var properties = new List<MapiProperty>();
        long decodedBytes = 0;
        long maximum = Math.Min(_options.MaxDecodedPropertyBytesPerItem,
            maximumDecodedBytes ?? _options.MaxDecodedPropertyBytesPerItem);
        if (rootHid != 0) {
            foreach (byte[] record in _heap.EnumerateBthLeafRecords(rootHid, 2, 6, indexLevels)) {
                _cancellationToken.ThrowIfCancellationRequested();
                ushort id = PstBinary.UInt16(record, 0);
                if (includedPropertyIds != null && !includedPropertyIds.Contains(id)) continue;
                if (properties.Count >= _options.MaxPropertiesPerItem) {
                    throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                        properties.Count + 1L, _options.MaxPropertiesPerItem);
                }
                var type = (MapiPropertyType)PstBinary.UInt16(record, 2);
                uint rawValue = PstBinary.UInt32(record, 4);
                if (deferredPropertyIds != null && deferredPropertyIds.Contains(id)) {
                    if (sourceHnids != null) sourceHnids[id] = rawValue;
                    properties.Add(new MapiProperty(id, type, null));
                    continue;
                }
                MapiProperty property = DecodeProperty(
                    id, type, rawValue, ref decodedBytes, maximum, sourceHnids);
                properties.Add(property);
            }
        }

        int codePage = ResolveCodePage(properties);
        foreach (MapiProperty property in properties) {
            if (property.PropertyType == MapiPropertyType.String8 && property.RawData != null) {
                property.Value = DecodeString8(property.RawData, codePage);
            }
        }
        return properties;
    }

    private MapiProperty DecodeProperty(ushort id, MapiPropertyType type, uint rawValue,
        ref long decodedBytes, long maximumDecodedBytes,
        IDictionary<ushort, uint>? sourceHnids) {
        object? value;
        byte[]? rawData = null;
        switch (type) {
            case MapiPropertyType.Null:
                value = null;
                break;
            case MapiPropertyType.Integer16:
                value = unchecked((short)(rawValue & 0xFFFF));
                break;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                value = unchecked((int)rawValue);
                break;
            case MapiPropertyType.Floating32:
                value = BitConverter.ToSingle(BitConverter.GetBytes(rawValue), 0);
                break;
            case MapiPropertyType.Boolean:
                value = (rawValue & 0xFF) != 0;
                break;
            default:
                if (sourceHnids != null) sourceHnids[id] = rawValue;
                rawData = _heap.ResolveHnid(rawValue, maximumDecodedBytes);
                decodedBytes = checked(decodedBytes + rawData.Length);
                if (decodedBytes > maximumDecodedBytes) {
                    throw new EmailStoreLimitExceededException(
                        nameof(EmailStoreItemReadOptions.MaxDecodedPropertyBytes), decodedBytes,
                        maximumDecodedBytes);
                }
                value = DecodeVariable(type, rawData);
                break;
        }
        return new MapiProperty(id, type, value) { RawData = rawData };
    }

    internal static object? DecodeVariable(MapiPropertyType type, byte[] bytes) {
        switch (type) {
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
                return bytes.Length >= 8 ? BitConverter.ToDouble(bytes, 0) : null;
            case MapiPropertyType.Currency:
            case MapiPropertyType.Integer64:
                return bytes.Length >= 8 ? (object)PstBinary.Int64(bytes, 0) : null;
            case MapiPropertyType.Time:
                if (bytes.Length < 8) return null;
                try {
                    return new DateTimeOffset(DateTime.FromFileTimeUtc(PstBinary.Int64(bytes, 0)));
                } catch (ArgumentOutOfRangeException) {
                    return null;
                }
            case MapiPropertyType.Guid:
                return bytes.Length >= 16 ? new Guid(bytes.Take(16).ToArray()) : (object?)null;
            case MapiPropertyType.Unicode:
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            case MapiPropertyType.String8:
                return DecodeString8(bytes, 1252);
            case MapiPropertyType.MultipleInteger16:
                return ReadInt16Array(bytes);
            case MapiPropertyType.MultipleInteger32:
                return ReadInt32Array(bytes);
            case MapiPropertyType.MultipleInteger64:
            case MapiPropertyType.MultipleTime:
                return ReadInt64Array(bytes);
            default:
                return bytes;
        }
    }

    private static short[] ReadInt16Array(byte[] bytes) {
        var values = new short[bytes.Length / 2];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int16(bytes, index * 2);
        return values;
    }

    private static int[] ReadInt32Array(byte[] bytes) {
        var values = new int[bytes.Length / 4];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int32(bytes, index * 4);
        return values;
    }

    private static long[] ReadInt64Array(byte[] bytes) {
        var values = new long[bytes.Length / 8];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int64(bytes, index * 8);
        return values;
    }

    private static int ResolveCodePage(IEnumerable<MapiProperty> properties) {
        foreach (ushort id in new ushort[] { 0x3FDE, 0x3FFC, 0x3FFD }) {
            MapiProperty? property = properties.FirstOrDefault(item => item.PropertyId == id);
            if (property?.Value is int value && value > 0 && value != 1200 && value != 1201) return value;
        }
        return 1252;
    }

    internal static string DecodeString8(byte[] bytes, int codePage) {
        try {
            return Encoding.GetEncoding(codePage).GetString(bytes).TrimEnd('\0');
        } catch (ArgumentException) {
            return Encoding.GetEncoding(1252).GetString(bytes).TrimEnd('\0');
        } catch (NotSupportedException) {
            return Encoding.GetEncoding(1252).GetString(bytes).TrimEnd('\0');
        }
    }
}
