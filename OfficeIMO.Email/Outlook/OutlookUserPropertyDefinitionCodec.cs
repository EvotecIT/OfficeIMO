namespace OfficeIMO.Email;

internal static class OutlookUserPropertyDefinitionCodec {
    internal const ushort Version1 = 0x0102;
    internal const ushort Version2 = 0x0103;
    private const int MaximumDefinitions = 16384;
    private const int MaximumSkipBlocksPerDefinition = 1024;

    static OutlookUserPropertyDefinitionCodec() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static ParseResult Parse(byte[]? bytes, int codePage) {
        if (bytes == null || bytes.Length == 0) return ParseResult.Missing;
        if (bytes.Length < 6) return ParseResult.Failed(OutlookUserPropertyDefinitionState.Corrupt,
            "The PropertyDefinition stream is shorter than its six-byte header.");

        var cursor = new Cursor(bytes);
        ushort version = cursor.ReadUInt16();
        if (version != Version1 && version != Version2) {
            return ParseResult.Failed(OutlookUserPropertyDefinitionState.UnsupportedVersion,
                string.Concat("Unsupported PropertyDefinition version 0x",
                    version.ToString("X4", CultureInfo.InvariantCulture), "."), version);
        }

        uint declaredCount = cursor.ReadUInt32();
        if (declaredCount > MaximumDefinitions) {
            return ParseResult.Failed(OutlookUserPropertyDefinitionState.Corrupt,
                "The PropertyDefinition stream declares too many field definitions.", version);
        }

        Encoding ansi;
        try {
            ansi = Encoding.GetEncoding(codePage, EncoderFallback.ReplacementFallback,
                DecoderFallback.ReplacementFallback);
        } catch (ArgumentException) {
            ansi = Encoding.GetEncoding(1252, EncoderFallback.ReplacementFallback,
                DecoderFallback.ReplacementFallback);
        }

        try {
            var definitions = new List<OutlookUserPropertyDefinition>((int)declaredCount);
            for (uint index = 0; index < declaredCount; index++) {
                int start = cursor.Offset;
                uint flags = cursor.ReadUInt32();
                ushort variantType = cursor.ReadUInt16();
                uint dispatchId = cursor.ReadUInt32();
                ushort unicodeNameLength = cursor.ReadUInt16();
                string unicodeName = cursor.ReadUnicode(unicodeNameLength);
                string ansiName = cursor.ReadPackedAnsi(ansi);
                string formula = cursor.ReadPackedAnsi(ansi);
                string validationRule = cursor.ReadPackedAnsi(ansi);
                string validationText = cursor.ReadPackedAnsi(ansi);
                cursor.ReadPackedAnsi(ansi); // ErrorANSI is reserved and not used by Outlook.

                OutlookUserPropertyType fieldType = InferFieldType(variantType);
                if (version == Version2) {
                    uint internalType = cursor.ReadUInt32();
                    fieldType = Enum.IsDefined(typeof(OutlookUserPropertyType), unchecked((int)internalType))
                        ? (OutlookUserPropertyType)unchecked((int)internalType)
                        : OutlookUserPropertyType.Unknown;

                    bool terminated = false;
                    for (int blockIndex = 0; blockIndex < MaximumSkipBlocksPerDefinition; blockIndex++) {
                        uint blockSize = cursor.ReadUInt32();
                        if (blockSize == 0) {
                            terminated = true;
                            break;
                        }
                        cursor.Skip(blockSize);
                    }
                    if (!terminated) throw new InvalidDataException("A V2 field definition has no terminating SkipBlock.");
                }

                string name = unicodeName.Length == 0 ? ansiName : unicodeName;
                if (string.IsNullOrEmpty(name)) {
                    throw new InvalidDataException("A field definition has no name.");
                }
                byte[] raw = cursor.Slice(start, cursor.Offset - start);
                definitions.Add(new OutlookUserPropertyDefinition(name, flags, variantType, dispatchId,
                    fieldType, formula, validationRule, validationText, version == Version2, raw));
            }

            if (!cursor.AtEnd) {
                throw new InvalidDataException("The PropertyDefinition stream contains undeclared trailing data.");
            }
            return ParseResult.Valid(version, definitions);
        } catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentOutOfRangeException ||
                                      ex is OverflowException) {
            return ParseResult.Failed(OutlookUserPropertyDefinitionState.Corrupt, ex.Message, version);
        }
    }

    internal static byte[] AddOrReplace(ParseResult parsed, string name, OutlookUserPropertyType fieldType,
        int codePage) {
        EnsureWritable(parsed);
        ushort version = parsed.State == OutlookUserPropertyDefinitionState.Missing ? Version2 : parsed.Version;
        var definitions = parsed.Definitions
            .Where(definition => !definition.IsCustom ||
                                 !string.Equals(definition.Name, name, StringComparison.OrdinalIgnoreCase))
            .Select(definition => definition.RawDefinition)
            .ToList();
        definitions.Add(EncodeDefinition(name, fieldType, version, codePage));
        return EncodeStream(version, definitions);
    }

    internal static byte[]? Remove(ParseResult parsed, string name) {
        if (parsed.State == OutlookUserPropertyDefinitionState.Missing) return null;
        EnsureWritable(parsed);
        var definitions = parsed.Definitions
            .Where(definition => !definition.IsCustom ||
                                 !string.Equals(definition.Name, name, StringComparison.OrdinalIgnoreCase))
            .Select(definition => definition.RawDefinition)
            .ToList();
        if (definitions.Count == parsed.Definitions.Count) return null;
        return definitions.Count == 0 ? Array.Empty<byte>() : EncodeStream(parsed.Version, definitions);
    }

    private static void EnsureWritable(ParseResult parsed) {
        if (parsed.State == OutlookUserPropertyDefinitionState.Valid ||
            parsed.State == OutlookUserPropertyDefinitionState.Missing) return;
        throw new InvalidOperationException(string.Concat(
            "The existing PropertyDefinition stream cannot be safely rewritten: ", parsed.Error));
    }

    private static byte[] EncodeStream(ushort version, IReadOnlyList<byte[]> definitions) {
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, true)) {
            writer.Write(version);
            writer.Write((uint)definitions.Count);
            foreach (byte[] definition in definitions) writer.Write(definition);
            writer.Flush();
            return stream.ToArray();
        }
    }

    private static byte[] EncodeDefinition(string name, OutlookUserPropertyType fieldType, ushort version,
        int codePage) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (name.Length == 0) throw new ArgumentException("A user property requires a name.", nameof(name));
        if (name.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(name));
        if (!CanCreate(fieldType)) {
            throw new ArgumentOutOfRangeException(nameof(fieldType), "That Outlook field type cannot be created directly.");
        }

        Encoding ansi;
        try {
            ansi = Encoding.GetEncoding(codePage, EncoderFallback.ReplacementFallback,
                DecoderFallback.ReplacementFallback);
        } catch (ArgumentException) {
            ansi = Encoding.GetEncoding(1252, EncoderFallback.ReplacementFallback,
                DecoderFallback.ReplacementFallback);
        }

        byte[] unicodeName = StrictUnicode().GetBytes(name);
        byte[] ansiName = ansi.GetBytes(name);
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, true)) {
            writer.Write(0x00000045U); // custom, print/save-as, and default print/save-as behavior
            writer.Write(GetVariantType(fieldType));
            writer.Write(0U);
            writer.Write(checked((ushort)name.Length));
            writer.Write(unicodeName);
            WritePackedBytes(writer, ansiName);
            WritePackedBytes(writer, Array.Empty<byte>()); // FormulaANSI
            WritePackedBytes(writer, Array.Empty<byte>()); // ValidationRuleANSI
            WritePackedBytes(writer, Array.Empty<byte>()); // ValidationTextANSI
            WritePackedBytes(writer, Array.Empty<byte>()); // ErrorANSI

            if (version == Version2) {
                writer.Write((uint)fieldType);
                using (var firstBlock = new MemoryStream())
                using (var firstWriter = new BinaryWriter(firstBlock, Encoding.UTF8, true)) {
                    WritePackedUnicode(firstWriter, name);
                    firstWriter.Flush();
                    byte[] content = firstBlock.ToArray();
                    writer.Write((uint)content.Length);
                    writer.Write(content);
                }
                writer.Write(0U);
            }
            writer.Flush();
            return stream.ToArray();
        }
    }

    private static void WritePackedBytes(BinaryWriter writer, byte[] bytes) {
        if (bytes.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (bytes.Length < byte.MaxValue) {
            writer.Write((byte)bytes.Length);
        } else {
            writer.Write(byte.MaxValue);
            writer.Write((ushort)bytes.Length);
        }
        writer.Write(bytes);
    }

    private static void WritePackedUnicode(BinaryWriter writer, string value) {
        if (value.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
        if (value.Length < byte.MaxValue) {
            writer.Write((byte)value.Length);
        } else {
            writer.Write(byte.MaxValue);
            writer.Write((ushort)value.Length);
        }
        writer.Write(StrictUnicode().GetBytes(value));
    }

    private static UnicodeEncoding StrictUnicode() => new UnicodeEncoding(false, false, true);

    private static bool CanCreate(OutlookUserPropertyType fieldType) {
        return fieldType == OutlookUserPropertyType.Text || fieldType == OutlookUserPropertyType.Number ||
               fieldType == OutlookUserPropertyType.Percent || fieldType == OutlookUserPropertyType.Currency ||
               fieldType == OutlookUserPropertyType.Boolean || fieldType == OutlookUserPropertyType.DateTime ||
               fieldType == OutlookUserPropertyType.Duration || fieldType == OutlookUserPropertyType.Keywords ||
               fieldType == OutlookUserPropertyType.Integer;
    }

    private static ushort GetVariantType(OutlookUserPropertyType fieldType) {
        switch (fieldType) {
            case OutlookUserPropertyType.Text: return 0x0008; // VT_BSTR
            case OutlookUserPropertyType.Number:
            case OutlookUserPropertyType.Percent: return 0x0005; // VT_R8
            case OutlookUserPropertyType.Currency: return 0x0006; // VT_CY
            case OutlookUserPropertyType.Boolean: return 0x000B; // VT_BOOL
            case OutlookUserPropertyType.DateTime: return 0x0007; // VT_DATE
            case OutlookUserPropertyType.Duration:
            case OutlookUserPropertyType.Integer: return 0x0003; // VT_I4
            case OutlookUserPropertyType.Keywords: return 0x2008; // VT_ARRAY | VT_BSTR
            default: throw new ArgumentOutOfRangeException(nameof(fieldType));
        }
    }

    private static OutlookUserPropertyType InferFieldType(ushort variantType) {
        switch (variantType) {
            case 0x0008:
            case 0x001E:
            case 0x001F: return OutlookUserPropertyType.Text;
            case 0x0005: return OutlookUserPropertyType.Number;
            case 0x0006: return OutlookUserPropertyType.Currency;
            case 0x0007:
            case 0x0040: return OutlookUserPropertyType.DateTime;
            case 0x000B: return OutlookUserPropertyType.Boolean;
            case 0x0003: return OutlookUserPropertyType.Integer;
            case 0x2008:
            case 0x101E:
            case 0x101F: return OutlookUserPropertyType.Keywords;
            default: return OutlookUserPropertyType.Unknown;
        }
    }

    internal sealed class ParseResult {
        private ParseResult(OutlookUserPropertyDefinitionState state, ushort version,
            IReadOnlyList<OutlookUserPropertyDefinition> definitions, string? error) {
            State = state;
            Version = version;
            Definitions = definitions;
            Error = error;
        }

        internal static readonly ParseResult Missing = new ParseResult(
            OutlookUserPropertyDefinitionState.Missing, 0,
            Array.Empty<OutlookUserPropertyDefinition>(), null);

        internal OutlookUserPropertyDefinitionState State { get; }
        internal ushort Version { get; }
        internal IReadOnlyList<OutlookUserPropertyDefinition> Definitions { get; }
        internal string? Error { get; }

        internal static ParseResult Valid(ushort version,
            IReadOnlyList<OutlookUserPropertyDefinition> definitions) =>
            new ParseResult(OutlookUserPropertyDefinitionState.Valid, version, definitions, null);

        internal static ParseResult Failed(OutlookUserPropertyDefinitionState state, string error,
            ushort version = 0) =>
            new ParseResult(state, version, Array.Empty<OutlookUserPropertyDefinition>(), error);
    }

    private sealed class Cursor {
        private readonly byte[] _bytes;

        internal Cursor(byte[] bytes) { _bytes = bytes; }
        internal int Offset { get; private set; }
        internal bool AtEnd => Offset == _bytes.Length;

        internal ushort ReadUInt16() {
            Ensure(2);
            ushort value = (ushort)(_bytes[Offset] | (_bytes[Offset + 1] << 8));
            Offset += 2;
            return value;
        }

        internal uint ReadUInt32() {
            Ensure(4);
            uint value = (uint)(_bytes[Offset] | (_bytes[Offset + 1] << 8) |
                                (_bytes[Offset + 2] << 16) | (_bytes[Offset + 3] << 24));
            Offset += 4;
            return value;
        }

        internal string ReadUnicode(int characterCount) {
            int byteCount = checked(characterCount * 2);
            Ensure(byteCount);
            string value = StrictUnicode().GetString(_bytes, Offset, byteCount);
            Offset += byteCount;
            return value;
        }

        internal string ReadPackedAnsi(Encoding encoding) {
            int byteCount = ReadPackedLength();
            Ensure(byteCount);
            string value = encoding.GetString(_bytes, Offset, byteCount);
            Offset += byteCount;
            return value;
        }

        internal void Skip(uint byteCount) {
            if (byteCount > int.MaxValue) throw new InvalidDataException("A SkipBlock is too large.");
            Ensure((int)byteCount);
            Offset += (int)byteCount;
        }

        internal byte[] Slice(int offset, int count) {
            var result = new byte[count];
            Buffer.BlockCopy(_bytes, offset, result, 0, count);
            return result;
        }

        private int ReadPackedLength() {
            Ensure(1);
            int length = _bytes[Offset++];
            return length == byte.MaxValue ? ReadUInt16() : length;
        }

        private void Ensure(int count) {
            if (count < 0 || Offset > _bytes.Length - count) {
                throw new InvalidDataException("The PropertyDefinition stream is truncated.");
            }
        }
    }
}
