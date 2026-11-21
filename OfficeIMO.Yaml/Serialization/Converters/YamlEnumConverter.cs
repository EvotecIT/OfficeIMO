using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// Enums as their member names. A <see cref="FlagsAttribute"/> enum combining several members is written as one
/// comma-separated scalar (<c>Read, Write</c>), the spelling of <see cref="Enum.ToString()"/> and <see cref="Enum.Parse(Type, string)"/>
/// that other YAML libraries also read; reading also accepts a sequence of names. Names are read ignoring case, and
/// integers are accepted. A value no member names is written as its number.
/// </summary>
internal sealed class YamlEnumConverter<TEnum> : YamlConverter<TEnum> where TEnum : struct, Enum {
    public override bool IsScalar => true;

    private readonly Dictionary<string, TEnum> _byName;
    private readonly Dictionary<TEnum, string> _names;
    private readonly ulong[] _bits;
    private readonly string[] _bitNames;
    private readonly bool _flags;
    private readonly bool _signed;

    public YamlEnumConverter() {
#if NET5_0_OR_GREATER
        TEnum[] values = Enum.GetValues<TEnum>();
        string[] names = Enum.GetNames<TEnum>();
#else
        var values = (TEnum[])Enum.GetValues(typeof(TEnum));
        string[] names = Enum.GetNames(typeof(TEnum));
#endif
        _flags = typeof(TEnum).IsDefined(typeof(FlagsAttribute), inherit: false);
        TypeCode code = Type.GetTypeCode(typeof(TEnum));
        _signed = code is TypeCode.SByte or TypeCode.Int16 or TypeCode.Int32 or TypeCode.Int64;
        _byName = new Dictionary<string, TEnum>(names.Length, StringComparer.OrdinalIgnoreCase);
        _names = new Dictionary<TEnum, string>(names.Length);
        var bits = new List<ulong>(names.Length);
        var bitNames = new List<string>(names.Length);
        for (int i = 0; i < names.Length; i++) {
            if (!_byName.ContainsKey(names[i])) _byName.Add(names[i], values[i]);
            if (!_names.ContainsKey(values[i])) _names.Add(values[i], names[i]);
            ulong raw = ToBits(values[i]);
            if (raw != 0) {
                bits.Add(raw);
                bitNames.Add(names[i]);
            }
        }
        _bits = bits.ToArray();
        _bitNames = bitNames.ToArray();
    }

    public override TEnum Read(YamlEventCursor cursor, YamlReadState state) {
        if (_flags && cursor.TokenType == YamlTokenType.SequenceStart) {
            ulong combined = 0;
            while (cursor.Read() && cursor.TokenType != YamlTokenType.SequenceEnd) {
                if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
                combined |= ToBits(ReadOne(cursor, state, YamlScalarText.Read(cursor, state, typeof(TEnum)).Trim()));
            }
            return FromBits(combined);
        }
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(TEnum)).Trim();
        if (!_flags || text.IndexOf(',') < 0) return ReadOne(cursor, state, text);
        ulong value = 0;
        while (true) {
            int comma = text.IndexOf(',');
            ReadOnlySpan<char> part = (comma < 0 ? text : text.Slice(0, comma)).Trim();
            value |= ToBits(ReadOne(cursor, state, part));
            if (comma < 0) break;
            text = text.Slice(comma + 1);
        }
        return FromBits(value);
    }

    private TEnum ReadOne(YamlEventCursor cursor, YamlReadState state, ReadOnlySpan<char> text) {
        if (text.Length > 0 && (char.IsDigit(text[0]) || text[0] == '-' || text[0] == '+')) {
            if (YamlSchemaResolver.TryParseInt64(text, cursor.Schema, out long number)) return FromBits(unchecked((ulong)number));
        } else if (_byName.TryGetValue(text.ToString(), out TEnum value)) {
            return value;
        }
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TEnum), "not a member of the enum");
    }

    public override void Write(YamlWriter writer, TEnum value, YamlWriteState state) {
        if (_names.TryGetValue(value, out string? name)) {
            writer.WriteScalar(name);
            return;
        }
        ulong raw = ToBits(value);
        if (_flags && raw != 0) {
            // Largest members first, as Enum.ToString does, so composite members win over their parts.
            var text = new StringBuilder();
            ulong remaining = raw;
            var used = new List<int>();
            for (int i = _bits.Length - 1; i >= 0; i--) {
                ulong bits = _bits[i];
                if ((remaining & bits) == bits) {
                    used.Add(i);
                    remaining &= ~bits;
                    if (remaining == 0) break;
                }
            }
            if (remaining == 0) {
                for (int i = used.Count - 1; i >= 0; i--) {
                    if (text.Length > 0) text.Append(", ");
                    text.Append(_bitNames[used[i]]);
                }
                writer.WriteScalar(text.ToString());
                return;
            }
        }
        if (_signed) writer.WriteScalar(unchecked((long)SignExtend(raw)));
        else if (raw <= long.MaxValue) writer.WriteScalar((long)raw);
        else YamlScalarText.WriteInvariant(writer, raw, null);
    }

    private static ulong SignExtend(ulong raw) => Unsafe.SizeOf<TEnum>() switch {
        1 => unchecked((ulong)(long)(sbyte)(byte)raw),
        2 => unchecked((ulong)(long)(short)(ushort)raw),
        4 => unchecked((ulong)(long)(int)(uint)raw),
        _ => raw
    };

    private static ulong ToBits(TEnum value) {
        switch (Unsafe.SizeOf<TEnum>()) {
            case 1: return Unsafe.As<TEnum, byte>(ref value);
            case 2: return Unsafe.As<TEnum, ushort>(ref value);
            case 4: return Unsafe.As<TEnum, uint>(ref value);
            default: return Unsafe.As<TEnum, ulong>(ref value);
        }
    }

    private static TEnum FromBits(ulong bits) {
        TEnum value = default;
        switch (Unsafe.SizeOf<TEnum>()) {
            case 1: Unsafe.As<TEnum, byte>(ref value) = (byte)bits; break;
            case 2: Unsafe.As<TEnum, ushort>(ref value) = (ushort)bits; break;
            case 4: Unsafe.As<TEnum, uint>(ref value) = (uint)bits; break;
            default: Unsafe.As<TEnum, ulong>(ref value) = bits; break;
        }
        return value;
    }
}
