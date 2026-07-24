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
            int scannedPropertyCount = 0;
            foreach (byte[] record in _heap.EnumerateBthLeafRecords(rootHid, 2, 6, indexLevels)) {
                _cancellationToken.ThrowIfCancellationRequested();
                scannedPropertyCount++;
                if (scannedPropertyCount > _options.MaxPropertiesPerItem) {
                    throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                        scannedPropertyCount, _options.MaxPropertiesPerItem);
                }
                ushort id = PstBinary.UInt16(record, 0);
                if (includedPropertyIds != null && !includedPropertyIds.Contains(id)) continue;
                var type = (MapiPropertyType)PstBinary.UInt16(record, 2);
                uint rawValue = PstBinary.UInt32(record, 4);
                if (deferredPropertyIds != null && deferredPropertyIds.Contains(id)) {
                    uint sourceHnid = rawValue;
                    byte[]? deferredRawData = null;
                    if (type == MapiPropertyType.Object && rawValue != 0 &&
                        (rawValue & 0x1F) == 0) {
                        // PtypObject stores an eight-byte (NID, size) descriptor in
                        // the heap. Resolving only that descriptor keeps the actual
                        // embedded object deferred while exposing its subnode NID.
                        deferredRawData = _heap.ResolveHnid(rawValue, 8);
                        if (deferredRawData.Length >= 8) {
                            sourceHnid = PstBinary.UInt32(deferredRawData, 0);
                        }
                    }
                    if (sourceHnids != null) sourceHnids[id] = sourceHnid;
                    properties.Add(new MapiProperty(id, type, null) { RawData = deferredRawData });
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
            } else if (property.PropertyType == MapiPropertyType.MultipleString8 && property.RawData != null) {
                property.Value = DecodeVariableElements(property.RawData)
                    .Select(value => DecodeString8(value, codePage))
                    .ToArray();
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
                rawData = _heap.ResolveHnid(rawValue, maximumDecodedBytes);
                decodedBytes = checked(decodedBytes + rawData.Length);
                if (decodedBytes > maximumDecodedBytes) {
                    throw new EmailStoreLimitExceededException(
                        nameof(EmailStoreItemReadOptions.MaxDecodedPropertyBytes), decodedBytes,
                        maximumDecodedBytes);
                }
                if (sourceHnids != null) {
                    uint sourceHnid = rawValue;
                    // A PC PtypObject value is an HNID whose allocation contains the
                    // referenced subnode NID followed by its declared size. Retain the
                    // historical direct-NID shape for compact third-party fixtures.
                    if (type == MapiPropertyType.Object && (rawValue & 0x1F) == 0 && rawData.Length >= 8) {
                        sourceHnid = PstBinary.UInt32(rawData, 0);
                    }
                    sourceHnids[id] = sourceHnid;
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
                return bytes.Length >= 8 ? (object)(PstBinary.Int64(bytes, 0) / 10000m) : null;
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
            case MapiPropertyType.MultipleFloating32:
                return ReadSingleArray(bytes);
            case MapiPropertyType.MultipleFloating64:
            case MapiPropertyType.MultipleFloatingTime:
                return ReadDoubleArray(bytes);
            case MapiPropertyType.MultipleCurrency:
                return ReadCurrencyArray(bytes);
            case MapiPropertyType.MultipleInteger64:
                return ReadInt64Array(bytes);
            case MapiPropertyType.MultipleTime:
                return ReadTimeArray(bytes);
            case MapiPropertyType.MultipleGuid:
                return ReadGuidArray(bytes);
            case MapiPropertyType.MultipleUnicode:
                return DecodeVariableElements(bytes)
                    .Select(value => Encoding.Unicode.GetString(value).TrimEnd('\0'))
                    .ToArray();
            case MapiPropertyType.MultipleString8:
                return DecodeVariableElements(bytes)
                    .Select(value => DecodeString8(value, 1252))
                    .ToArray();
            case MapiPropertyType.MultipleBinary:
                return DecodeVariableElements(bytes);
            default:
                return bytes;
        }
    }

    private static short[] ReadInt16Array(byte[] bytes) {
        EnsurePackedElementSize(bytes, 2);
        var values = new short[bytes.Length / 2];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int16(bytes, index * 2);
        return values;
    }

    private static int[] ReadInt32Array(byte[] bytes) {
        EnsurePackedElementSize(bytes, 4);
        var values = new int[bytes.Length / 4];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int32(bytes, index * 4);
        return values;
    }

    private static long[] ReadInt64Array(byte[] bytes) {
        EnsurePackedElementSize(bytes, 8);
        var values = new long[bytes.Length / 8];
        for (int index = 0; index < values.Length; index++) values[index] = PstBinary.Int64(bytes, index * 8);
        return values;
    }

    private static decimal[] ReadCurrencyArray(byte[] bytes) {
        EnsurePackedElementSize(bytes, 8);
        var values = new decimal[bytes.Length / 8];
        for (int index = 0; index < values.Length; index++)
            values[index] = PstBinary.Int64(bytes, index * 8) / 10000m;
        return values;
    }

    private static float[] ReadSingleArray(byte[] bytes) {
        EnsurePackedElementSize(bytes, 4);
        var values = new float[bytes.Length / 4];
        for (int index = 0; index < values.Length; index++) values[index] = BitConverter.ToSingle(bytes, index * 4);
        return values;
    }

    private static double[] ReadDoubleArray(byte[] bytes) {
        EnsurePackedElementSize(bytes, 8);
        var values = new double[bytes.Length / 8];
        for (int index = 0; index < values.Length; index++) values[index] = BitConverter.ToDouble(bytes, index * 8);
        return values;
    }

    private static DateTimeOffset?[] ReadTimeArray(byte[] bytes) {
        EnsurePackedElementSize(bytes, 8);
        var values = new DateTimeOffset?[bytes.Length / 8];
        for (int index = 0; index < values.Length; index++) {
            try {
                values[index] = new DateTimeOffset(DateTime.FromFileTimeUtc(PstBinary.Int64(bytes, index * 8)));
            } catch (ArgumentOutOfRangeException) {
                values[index] = null;
            }
        }
        return values;
    }

    private static Guid[] ReadGuidArray(byte[] bytes) {
        EnsurePackedElementSize(bytes, 16);
        var values = new Guid[bytes.Length / 16];
        for (int index = 0; index < values.Length; index++) {
            var value = new byte[16];
            Buffer.BlockCopy(bytes, index * 16, value, 0, value.Length);
            values[index] = new Guid(value);
        }
        return values;
    }

    private static byte[][] DecodeVariableElements(byte[] bytes) {
        if (bytes.Length < 4) throw new InvalidDataException("A variable multi-valued PST property is truncated.");
        uint rawCount = PstBinary.UInt32(bytes, 0);
        if (rawCount > int.MaxValue) {
            throw new InvalidDataException("A variable multi-valued PST property declares too many elements.");
        }
        int count = (int)rawCount;
        int dataStart = checked(4 + count * 4);
        if (dataStart > bytes.Length) {
            throw new InvalidDataException("A variable multi-valued PST property offset table is truncated.");
        }
        var values = new byte[count][];
        int previous = dataStart;
        for (int index = 0; index < count; index++) {
            int start = ReadBoundedOffset(bytes, 4 + index * 4);
            int end = index + 1 < count
                ? ReadBoundedOffset(bytes, 4 + (index + 1) * 4)
                : bytes.Length;
            if (start < dataStart || start < previous || end < start || end > bytes.Length) {
                throw new InvalidDataException("A variable multi-valued PST property contains an invalid element offset.");
            }
            var value = new byte[end - start];
            Buffer.BlockCopy(bytes, start, value, 0, value.Length);
            values[index] = value;
            previous = start;
        }
        return values;
    }

    private static int ReadBoundedOffset(byte[] bytes, int offset) {
        uint value = PstBinary.UInt32(bytes, offset);
        if (value > int.MaxValue) {
            throw new InvalidDataException("A variable multi-valued PST property offset is out of range.");
        }
        return (int)value;
    }

    private static void EnsurePackedElementSize(byte[] bytes, int elementSize) {
        if (bytes.Length % elementSize != 0) {
            throw new InvalidDataException("A fixed multi-valued PST property has a partial trailing element.");
        }
    }

    private static int ResolveCodePage(IEnumerable<MapiProperty> properties) {
        foreach (MapiPropertyKey<int> key in new[] {
            MapiKnownProperties.PidTag.InternetCodepage,
            MapiKnownProperties.PidTag.CodePageId,
            MapiKnownProperties.PidTag.MessageCodepage
        }) {
            MapiProperty? property = properties.FirstOrDefault(key.Matches);
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
