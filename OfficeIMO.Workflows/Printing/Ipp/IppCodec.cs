namespace OfficeIMO.Workflows;

/// <summary>Bounds applied when decoding a message received from a printer.</summary>
internal sealed class IppDecodeLimits {
    public static readonly IppDecodeLimits Default = new();

    /// <summary>Largest accepted message, including any trailing data.</summary>
    public int MaxMessageBytes { get; init; } = 8 * 1024 * 1024;
    /// <summary>Attribute groups, including repeated job groups in Get-Jobs responses.</summary>
    public int MaxGroups { get; init; } = 1024;
    /// <summary>Attributes and collection members, counted together.</summary>
    public int MaxAttributes { get; init; } = 32768;
    /// <summary>Values across all attributes and collection members.</summary>
    public int MaxValues { get; init; } = 262144;
    /// <summary>Nesting of collection values; media-col needs two levels.</summary>
    public int MaxCollectionDepth { get; init; } = 8;
}

/// <summary>Binary IPP encoding (RFC 8010 section 3).</summary>
internal static class IppCodec {
    private const int MaxShort = short.MaxValue;

    public static byte[] Encode(IppMessage message) {
        using var output = new MemoryStream();
        Encode(message, output);
        return output.ToArray();
    }

    public static void Encode(IppMessage message, Stream output) {
        if (message == null) throw new ArgumentNullException(nameof(message));
        if (output == null) throw new ArgumentNullException(nameof(output));
        var writer = new Writer(output);
        writer.Byte(message.VersionMajor);
        writer.Byte(message.VersionMinor);
        writer.Short(message.Code);
        writer.Int(message.RequestId);
        foreach (IppAttributeGroup group in message.Groups) {
            if ((byte)group.Tag is 0 or >= 0x10 || group.Tag == IppGroupTag.EndOfAttributes) throw new ArgumentException("Invalid attribute group tag.", nameof(message));
            writer.Byte((byte)group.Tag);
            foreach (IppAttribute attribute in group.Attributes) WriteAttribute(writer, attribute, 0);
        }
        writer.Byte((byte)IppGroupTag.EndOfAttributes);
    }

    private static void WriteAttribute(Writer writer, IppAttribute attribute, int depth) {
        if (attribute.Name.Length == 0) throw new ArgumentException("Attribute names cannot be empty.");
        if (attribute.Values.Count == 0) throw new ArgumentException($"Attribute '{attribute.Name}' has no values.");
        byte[] name = IppValue.EncodeString(attribute.Name);
        for (int i = 0; i < attribute.Values.Count; i++) WriteValue(writer, i == 0 ? name : Array.Empty<byte>(), attribute.Values[i], depth);
    }

    private static void WriteValue(Writer writer, byte[] name, IppValue value, int depth) {
        if (value.Tag == IppValueTag.BegCollection) {
            if (depth >= IppDecodeLimits.Default.MaxCollectionDepth) throw new ArgumentException("Collections are nested too deeply.");
            writer.Header(IppValueTag.BegCollection, name, 0);
            foreach (IppAttribute member in value.Members!) {
                if (member.Values.Count == 0) throw new ArgumentException($"Member '{member.Name}' has no values.");
                writer.Header(IppValueTag.MemberAttrName, Array.Empty<byte>(), Checked(IppValue.EncodeString(member.Name)));
                foreach (IppValue memberValue in member.Values) WriteValue(writer, Array.Empty<byte>(), memberValue, depth + 1);
            }
            writer.Header(IppValueTag.EndCollection, Array.Empty<byte>(), 0);
            return;
        }
        switch (value.Tag) {
            case IppValueTag.Integer or IppValueTag.Enum:
                writer.Header(value.Tag, name, 4);
                writer.Int(value.Integer);
                return;
            case IppValueTag.Boolean:
                writer.Header(value.Tag, name, 1);
                writer.Byte((byte)(value.Boolean ? 1 : 0));
                return;
            case IppValueTag.RangeOfInteger:
                writer.Header(value.Tag, name, 8);
                writer.Int(value.Integer);
                writer.Int(value.Second);
                return;
            case IppValueTag.Resolution:
                writer.Header(value.Tag, name, 9);
                writer.Int(value.Integer);
                writer.Int(value.Second);
                writer.Byte(value.Units);
                return;
            case IppValueTag.TextWithLanguage or IppValueTag.NameWithLanguage: {
                byte[] language = IppValue.EncodeString(value.Language!);
                byte[] text = IppValue.EncodeString(value.Text!);
                writer.Header(value.Tag, name, Checked(4 + language.Length + text.Length));
                writer.Short((short)language.Length);
                writer.Bytes(language);
                writer.Short((short)text.Length);
                writer.Bytes(text);
                return;
            }
            default:
                if (value.IsOutOfBand) {
                    writer.Header(value.Tag, name, 0);
                } else if (value.Text is not null) {
                    writer.Header(value.Tag, name, Checked(IppValue.EncodeString(value.Text)));
                } else {
                    writer.Header(value.Tag, name, Checked(value.Octets ?? Array.Empty<byte>()));
                }
                return;
        }
    }

    private static byte[] Checked(byte[] value) {
        if (value.Length > MaxShort) throw new ArgumentException("An IPP value cannot exceed 32767 bytes.");
        return value;
    }

    private static int Checked(int length) => length > MaxShort ? throw new ArgumentException("An IPP value cannot exceed 32767 bytes.") : length;

    /// <summary>Decodes an IPP message. Any malformed or over-limit input raises <see cref="IppException"/>.</summary>
    /// <param name="data">Message bytes.</param>
    /// <param name="length">Number of valid bytes in <paramref name="data"/>.</param>
    /// <param name="limits">Decoder bounds, or the defaults.</param>
    /// <param name="dataOffset">Offset of any document data that follows the end-of-attributes tag.</param>
    public static IppMessage Decode(byte[] data, int length, IppDecodeLimits? limits, out int dataOffset) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (length < 0 || length > data.Length) throw new ArgumentOutOfRangeException(nameof(length));
        limits ??= IppDecodeLimits.Default;
        if (length > limits.MaxMessageBytes) throw Malformed($"The IPP message exceeds {limits.MaxMessageBytes} bytes.");
        var reader = new Reader(data, length, limits);
        IppMessage message = reader.ReadMessage();
        dataOffset = reader.Position;
        return message;
    }

    public static IppMessage Decode(byte[] data, IppDecodeLimits? limits = null) => Decode(data, data?.Length ?? 0, limits, out _);

    internal static IppException Malformed(string message) => new("Malformed IPP message: " + message);

    private sealed class Writer {
        private readonly Stream _output;
        private readonly byte[] _buffer = new byte[4];
        public Writer(Stream output) => _output = output;
        public void Byte(byte value) => _output.WriteByte(value);
        public void Bytes(byte[] value) => _output.Write(value, 0, value.Length);
        public void Short(short value) {
            _buffer[0] = (byte)(value >> 8);
            _buffer[1] = (byte)value;
            _output.Write(_buffer, 0, 2);
        }
        public void Int(int value) {
            _buffer[0] = (byte)(value >> 24);
            _buffer[1] = (byte)(value >> 16);
            _buffer[2] = (byte)(value >> 8);
            _buffer[3] = (byte)value;
            _output.Write(_buffer, 0, 4);
        }
        public void Header(IppValueTag tag, byte[] name, int valueLength) {
            if (name.Length > MaxShort) throw new ArgumentException("An IPP attribute name cannot exceed 32767 bytes.");
            Byte((byte)tag);
            Short((short)name.Length);
            Bytes(name);
            Short((short)valueLength);
        }
        public void Header(IppValueTag tag, byte[] name, byte[] value) {
            Header(tag, name, value.Length);
            Bytes(value);
        }
    }

    private sealed class Reader {
        private readonly byte[] _data;
        private readonly int _length;
        private readonly IppDecodeLimits _limits;
        private int _attributes;
        private int _values;

        public Reader(byte[] data, int length, IppDecodeLimits limits) {
            _data = data;
            _length = length;
            _limits = limits;
        }

        public int Position { get; private set; }

        public IppMessage ReadMessage() {
            Require(8, "header");
            var message = new IppMessage(_data[0], _data[1], ReadShort(2), ReadInt(4));
            Position = 8;
            if (message.VersionMajor is 0 or > 9) throw Malformed($"unsupported version {message.VersionMajor}.{message.VersionMinor}.");
            IppAttributeGroup? group = null;
            IppAttribute? current = null;
            while (true) {
                Require(1, "tag");
                byte tag = _data[Position++];
                if (tag == (byte)IppGroupTag.EndOfAttributes) return message;
                if (tag < 0x10) {
                    if (tag == 0) throw Malformed("reserved delimiter tag 0x00.");
                    if (message.Groups.Count >= _limits.MaxGroups) throw Malformed($"more than {_limits.MaxGroups} attribute groups.");
                    group = message.AddGroup((IppGroupTag)tag);
                    current = null;
                    continue;
                }
                if (group is null) throw Malformed("an attribute appears before any attribute group.");
                string name = ReadName();
                if (tag is (byte)IppValueTag.EndCollection or (byte)IppValueTag.MemberAttrName) throw Malformed($"tag 0x{tag:X2} outside a collection.");
                IppValue value = ReadValue(tag, 0);
                if (name.Length > 0) {
                    current = NewAttribute(name);
                    group.Attributes.Add(current);
                } else if (current is null) {
                    throw Malformed("an additional value has no attribute.");
                }
                current.Values.Add(value);
            }
        }

        private IppAttribute NewAttribute(string name) {
            if (++_attributes > _limits.MaxAttributes) throw Malformed($"more than {_limits.MaxAttributes} attributes.");
            return new IppAttribute(name);
        }

        private string ReadName() {
            int length = ReadLength("name");
            string name = IppValue.DecodeString(_data, Position, length);
            Position += length;
            return name;
        }

        private IppValue ReadValue(byte tag, int depth) {
            if (++_values > _limits.MaxValues) throw Malformed($"more than {_limits.MaxValues} values.");
            int length = ReadLength("value");
            int start = Position;
            Position += length;
            if (tag == (byte)IppValueTag.BegCollection) return ReadCollection(depth + 1);
            return Decode((IppValueTag)tag, start, length);
        }

        private IppValue ReadCollection(int depth) {
            if (depth > _limits.MaxCollectionDepth) throw Malformed($"collections nested deeper than {_limits.MaxCollectionDepth}.");
            var members = new List<IppAttribute>();
            IppAttribute? member = null;
            while (true) {
                Require(1, "collection tag");
                byte tag = _data[Position++];
                if (tag < 0x10) throw Malformed("a collection is not terminated.");
                if (ReadLength("name") != 0) throw Malformed("a collection member value carries a name.");
                if (tag is (byte)IppValueTag.EndCollection or (byte)IppValueTag.MemberAttrName && member is { Values.Count: 0 })
                    throw Malformed($"collection member '{member.Name}' has no value.");
                if (tag == (byte)IppValueTag.EndCollection) {
                    int ignored = ReadLength("value");
                    Position += ignored;
                    return IppValue.OfCollection(members);
                }
                if (tag == (byte)IppValueTag.MemberAttrName) {
                    int length = ReadLength("member name");
                    if (length == 0) throw Malformed("an empty collection member name.");
                    member = NewAttribute(IppValue.DecodeString(_data, Position, length));
                    Position += length;
                    members.Add(member);
                    continue;
                }
                if (member is null) throw Malformed("a collection value has no member name.");
                member.Values.Add(ReadValue(tag, depth));
            }
        }

        private IppValue Decode(IppValueTag tag, int start, int length) {
            byte raw = (byte)tag;
            if (raw is >= 0x10 and <= 0x1F) return IppValue.OutOfBand(tag);
            switch (tag) {
                case IppValueTag.Integer or IppValueTag.Enum:
                    Exact(tag, length, 4);
                    return IppValue.Decoded(tag, ReadInt(start), 0, 0, null, null, null);
                case IppValueTag.Boolean:
                    Exact(tag, length, 1);
                    if (_data[start] > 1) throw Malformed($"boolean value {_data[start]}.");
                    return IppValue.Decoded(tag, _data[start], 0, 0, null, null, null);
                case IppValueTag.RangeOfInteger:
                    Exact(tag, length, 8);
                    return IppValue.Decoded(tag, ReadInt(start), ReadInt(start + 4), 0, null, null, null);
                case IppValueTag.Resolution:
                    Exact(tag, length, 9);
                    return IppValue.Decoded(tag, ReadInt(start), ReadInt(start + 4), _data[start + 8], null, null, null);
                case IppValueTag.DateTime:
                    Exact(tag, length, 11);
                    return IppValue.Decoded(tag, 0, 0, 0, null, null, Copy(start, length));
                case IppValueTag.TextWithLanguage or IppValueTag.NameWithLanguage: {
                    if (length < 4) throw Malformed($"{tag} value is too short.");
                    int languageLength = (ushort)ReadShort(start);
                    if (2 + languageLength + 2 > length) throw Malformed($"{tag} language overruns its value.");
                    int textLength = (ushort)ReadShort(start + 2 + languageLength);
                    if (4 + languageLength + textLength != length) throw Malformed($"{tag} lengths do not match its value.");
                    return IppValue.Decoded(tag, 0, 0, 0,
                        IppValue.DecodeString(_data, start + 4 + languageLength, textLength),
                        IppValue.DecodeString(_data, start + 2, languageLength), null);
                }
                case IppValueTag.Extension:
                    if (length < 4) throw Malformed("extension value has no extended tag.");
                    return IppValue.Opaque(raw, Copy(start, length));
            }
            if (raw is >= 0x40 and <= 0x5F) return IppValue.Decoded(tag, 0, 0, 0, IppValue.DecodeString(_data, start, length), null, null);
            return IppValue.Opaque(raw, Copy(start, length));
        }

        private static void Exact(IppValueTag tag, int length, int expected) {
            if (length != expected) throw Malformed($"{tag} value has {length} bytes instead of {expected}.");
        }

        private byte[] Copy(int start, int length) {
            var copy = new byte[length];
            Buffer.BlockCopy(_data, start, copy, 0, length);
            return copy;
        }

        private int ReadLength(string what) {
            Require(2, what + " length");
            short length = ReadShort(Position);
            Position += 2;
            if (length < 0) throw Malformed($"negative {what} length.");
            Require(length, what);
            return length;
        }

        private void Require(int count, string what) {
            if (count > _length - Position) throw Malformed($"truncated {what}.");
        }

        private short ReadShort(int at) => (short)((_data[at] << 8) | _data[at + 1]);
        private int ReadInt(int at) => (_data[at] << 24) | (_data[at + 1] << 16) | (_data[at + 2] << 8) | _data[at + 3];
    }
}
