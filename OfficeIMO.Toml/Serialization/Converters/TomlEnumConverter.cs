using System.Runtime.CompilerServices;

namespace OfficeIMO.Toml;

/// <summary>
/// Enums as their member names, as YAML writes them. A <see cref="FlagsAttribute"/> enum combining several members is
/// one comma-separated string (<c>"Read, Write"</c>); reading also accepts an array of names. Names are read ignoring
/// case, and integers are accepted. A value no member names is written as its number.
/// </summary>
internal sealed class TomlEnumConverter<TEnum> : TomlConverter<TEnum> where TEnum : struct, Enum {
    private readonly Dictionary<string, TEnum> _byName;
    private readonly Dictionary<TEnum, string> _names;
    private readonly ulong[] _bits;
    private readonly string[] _bitNames;
    private readonly bool _flags;
    private readonly bool _signed;

    public TomlEnumConverter() {
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

    public override TEnum Read(TomlNode node, TomlReadState state) {
        if (_flags && node is TomlArray array) {
            ulong combined = 0;
            foreach (TomlNode item in array) combined |= ToBits(ReadOne(item, Text(item)));
            return FromBits(combined);
        }
        if (node is TomlValue { Kind: TomlNodeKind.Integer } integer) return FromBits(unchecked((ulong)TomlValueText.GetInt64(integer)));
        string text = Text(node).Trim();
        if (!_flags || text.IndexOf(',') < 0) return ReadOne(node, text);
        ulong value = 0;
        foreach (string part in text.Split(',')) value |= ToBits(ReadOne(node, part.Trim()));
        return FromBits(value);
    }

    private static string Text(TomlNode node) =>
        node is TomlValue { Kind: TomlNodeKind.String } value ? TomlValueText.GetString(value) : throw TomlSerializationErrors.TypeMismatch(node, typeof(TEnum));

    private TEnum ReadOne(TomlNode node, string text) {
        if (_byName.TryGetValue(text, out TEnum value)) return value;
        throw TomlSerializationErrors.TypeMismatch(node, typeof(TEnum), "not a member of the enum");
    }

    public override bool TryReadKey(string key, out TEnum value) => _byName.TryGetValue(key, out value);

    public override string? WriteKey(TEnum value) => Format(value);

    public override void Write(TomlEmitter emitter, TEnum value, TomlWriteState state) {
        string? name = Format(value);
        if (name is not null) {
            emitter.WriteString(name);
            return;
        }
        ulong raw = ToBits(value);
        if (_signed) emitter.WriteInteger(unchecked((long)SignExtend(raw)));
        else if (raw <= long.MaxValue) emitter.WriteInteger((long)raw);
        else throw TomlSerializationErrors.OutOfRange(raw.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>The member name, or for flags the comma-separated names of the members that make up the value, or null.</summary>
    private string? Format(TEnum value) {
        if (_names.TryGetValue(value, out string? name)) return name;
        ulong raw = ToBits(value);
        if (!_flags || raw == 0) return null;
        // Largest members first, as Enum.ToString does, so composite members win over their parts.
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
        if (remaining != 0) return null;
        var text = new StringBuilder();
        for (int i = used.Count - 1; i >= 0; i--) {
            if (text.Length > 0) text.Append(", ");
            text.Append(_bitNames[used[i]]);
        }
        return text.ToString();
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
