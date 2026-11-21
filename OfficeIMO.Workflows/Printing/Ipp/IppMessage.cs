using System.Text;

namespace OfficeIMO.Workflows;

/// <summary>Delimiter tags that start an attribute group (RFC 8010 section 3.5.1).</summary>
internal enum IppGroupTag : byte {
    Operation = 0x01,
    Job = 0x02,
    EndOfAttributes = 0x03,
    Printer = 0x04,
    Unsupported = 0x05,
    Subscription = 0x06,
    EventNotification = 0x07,
    Resource = 0x08,
    Document = 0x09,
    System = 0x0A
}

/// <summary>Value tags (RFC 8010 section 3.5.2).</summary>
internal enum IppValueTag : byte {
    Unsupported = 0x10,
    Default = 0x11,
    Unknown = 0x12,
    NoValue = 0x13,
    NotSettable = 0x15,
    DeleteAttribute = 0x16,
    AdminDefine = 0x17,
    Integer = 0x21,
    Boolean = 0x22,
    Enum = 0x23,
    OctetString = 0x30,
    DateTime = 0x31,
    Resolution = 0x32,
    RangeOfInteger = 0x33,
    BegCollection = 0x34,
    TextWithLanguage = 0x35,
    NameWithLanguage = 0x36,
    EndCollection = 0x37,
    TextWithoutLanguage = 0x41,
    NameWithoutLanguage = 0x42,
    Keyword = 0x44,
    Uri = 0x45,
    UriScheme = 0x46,
    Charset = 0x47,
    NaturalLanguage = 0x48,
    MimeMediaType = 0x49,
    MemberAttrName = 0x4A,
    Extension = 0x7F
}

/// <summary>Operation identifiers used by the client (RFC 8011 section 5.4.15).</summary>
internal enum IppOperation : short {
    PrintJob = 0x0002,
    ValidateJob = 0x0004,
    CancelJob = 0x0008,
    GetJobAttributes = 0x0009,
    GetJobs = 0x000A,
    GetPrinterAttributes = 0x000B,
    CupsGetDefault = 0x4001,
    CupsGetPrinters = 0x4002
}

/// <summary>One attribute value. Which members are meaningful depends on <see cref="Tag"/>.</summary>
internal sealed class IppValue {
    private static readonly UTF8Encoding Utf8 = new(false, false);

    private IppValue(IppValueTag tag) => Tag = tag;

    public IppValueTag Tag { get; }
    /// <summary>integer, enum, boolean (0/1), resolution cross-feed, or range lower bound.</summary>
    public int Integer { get; private set; }
    /// <summary>Resolution feed direction or range upper bound.</summary>
    public int Second { get; private set; }
    /// <summary>Resolution units (3 = dots per inch, 4 = dots per centimetre).</summary>
    public byte Units { get; private set; }
    public string? Text { get; private set; }
    public string? Language { get; private set; }
    /// <summary>octetString, dateTime and unrecognized value bytes.</summary>
    public byte[]? Octets { get; private set; }
    public IReadOnlyList<IppAttribute>? Members { get; private set; }

    public bool IsOutOfBand => (byte)Tag is >= 0x10 and <= 0x1F;
    public bool Boolean => Integer != 0;

    public static IppValue OutOfBand(IppValueTag tag) {
        if ((byte)tag is < 0x10 or > 0x1F) throw new ArgumentOutOfRangeException(nameof(tag));
        return new(tag);
    }
    public static IppValue OfInteger(int value) => new(IppValueTag.Integer) { Integer = value };
    public static IppValue OfEnum(int value) => new(IppValueTag.Enum) { Integer = value };
    public static IppValue OfBoolean(bool value) => new(IppValueTag.Boolean) { Integer = value ? 1 : 0 };
    public static IppValue OfRange(int lower, int upper) => new(IppValueTag.RangeOfInteger) { Integer = lower, Second = upper };
    public static IppValue OfResolution(int crossFeed, int feed, byte units) => new(IppValueTag.Resolution) { Integer = crossFeed, Second = feed, Units = units };
    public static IppValue OfOctets(byte[] value) => new(IppValueTag.OctetString) { Octets = value ?? throw new ArgumentNullException(nameof(value)) };
    public static IppValue OfCollection(IReadOnlyList<IppAttribute> members) => new(IppValueTag.BegCollection) { Members = members ?? throw new ArgumentNullException(nameof(members)) };

    public static IppValue OfString(IppValueTag tag, string value) {
        if ((byte)tag is < 0x40 or > 0x5F || tag == IppValueTag.MemberAttrName) throw new ArgumentOutOfRangeException(nameof(tag));
        return new(tag) { Text = value ?? throw new ArgumentNullException(nameof(value)) };
    }

    public static IppValue OfStringWithLanguage(IppValueTag tag, string value, string language) {
        if (tag is not (IppValueTag.TextWithLanguage or IppValueTag.NameWithLanguage)) throw new ArgumentOutOfRangeException(nameof(tag));
        return new(tag) { Text = value ?? throw new ArgumentNullException(nameof(value)), Language = language ?? throw new ArgumentNullException(nameof(language)) };
    }

    public static IppValue OfDateTime(DateTimeOffset value) {
        TimeSpan offset = value.Offset;
        int minutes = (int)Math.Abs(offset.TotalMinutes);
        return new(IppValueTag.DateTime) {
            Octets = new[] {
                (byte)(value.Year >> 8), (byte)value.Year, (byte)value.Month, (byte)value.Day,
                (byte)value.Hour, (byte)value.Minute, (byte)value.Second, (byte)(value.Millisecond / 100),
                (byte)(offset < TimeSpan.Zero ? '-' : '+'), (byte)(minutes / 60), (byte)(minutes % 60)
            }
        };
    }

    /// <summary>Wraps bytes for a tag the codec keeps opaque (unknown types and the extension tag).</summary>
    public static IppValue Opaque(byte tag, byte[] value) => new((IppValueTag)tag) { Octets = value };

    internal static IppValue Decoded(IppValueTag tag, int integer, int second, byte units, string? text, string? language, byte[]? octets) =>
        new(tag) { Integer = integer, Second = second, Units = units, Text = text, Language = language, Octets = octets };

    /// <summary>Reads an RFC 2579 DateAndTime value; false when the bytes are not a valid instant.</summary>
    public bool TryGetDateTime(out DateTimeOffset value) {
        value = default;
        if (Tag != IppValueTag.DateTime || Octets is not { Length: 11 } b) return false;
        int year = (b[0] << 8) | b[1];
        if (year is < 1 or > 9999 || b[2] is < 1 or > 12 || b[3] < 1 || b[3] > DateTime.DaysInMonth(year, b[2]) ||
            b[4] > 23 || b[5] > 59 || b[6] > 60 || b[7] > 9 || (b[8] != '+' && b[8] != '-') || b[10] > 59 || b[9] * 60 + b[10] > 14 * 60) return false;
        var offset = new TimeSpan(b[9], b[10], 0);
        if (b[8] == '-') offset = offset.Negate();
        try {
            value = new DateTimeOffset(year, b[2], b[3], b[4], b[5], Math.Min(b[6], (byte)59), b[7] * 100, offset);
            return true;
        } catch (ArgumentOutOfRangeException) {
            // The local time is valid but its UTC instant falls outside DateTimeOffset's range.
            return false;
        }
    }

    internal static string DecodeString(byte[] data, int offset, int count) => Utf8.GetString(data, offset, count);
    internal static byte[] EncodeString(string value) => Utf8.GetBytes(value);

    public override string ToString() => Tag switch {
        IppValueTag.Integer or IppValueTag.Enum => Integer.ToString(System.Globalization.CultureInfo.InvariantCulture),
        IppValueTag.Boolean => Boolean ? "true" : "false",
        IppValueTag.RangeOfInteger => Integer + "-" + Second,
        IppValueTag.Resolution => Integer + "x" + Second + (Units == 3 ? "dpi" : "dpcm"),
        IppValueTag.BegCollection => "{" + string.Join(" ", Members!.Select(member => member.ToString())) + "}",
        _ when Text is not null => Text,
        _ when IsOutOfBand => Tag.ToString(),
        _ => "<" + (Octets?.Length ?? 0) + " bytes>"
    };
}

/// <summary>A named attribute with one or more values.</summary>
internal sealed class IppAttribute {
    public IppAttribute(string name, params IppValue[] values) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Values = new List<IppValue>(values ?? throw new ArgumentNullException(nameof(values)));
    }

    public string Name { get; }
    public List<IppValue> Values { get; }
    public IppValue? First => Values.Count == 0 ? null : Values[0];

    public override string ToString() => Name + "=" + string.Join(",", Values.Select(value => value.ToString()));
}

/// <summary>An attribute group introduced by a delimiter tag.</summary>
internal sealed class IppAttributeGroup {
    public IppAttributeGroup(IppGroupTag tag) => Tag = tag;
    public IppGroupTag Tag { get; }
    public List<IppAttribute> Attributes { get; } = new();

    public IppAttribute? this[string name] {
        get {
            foreach (IppAttribute attribute in Attributes) {
                if (string.Equals(attribute.Name, name, StringComparison.Ordinal)) return attribute;
            }
            return null;
        }
    }

    public IppAttributeGroup Add(string name, params IppValue[] values) {
        Attributes.Add(new IppAttribute(name, values));
        return this;
    }
}

/// <summary>An IPP request or response (RFC 8010 section 3.1.1).</summary>
internal sealed class IppMessage {
    public IppMessage(byte versionMajor, byte versionMinor, short code, int requestId) {
        VersionMajor = versionMajor;
        VersionMinor = versionMinor;
        Code = code;
        RequestId = requestId;
    }

    public byte VersionMajor { get; set; }
    public byte VersionMinor { get; set; }
    /// <summary>operation-id in requests, status-code in responses.</summary>
    public short Code { get; }
    public int RequestId { get; set; }
    public List<IppAttributeGroup> Groups { get; } = new();

    public IppAttributeGroup AddGroup(IppGroupTag tag) {
        var group = new IppAttributeGroup(tag);
        Groups.Add(group);
        return group;
    }

    public IppAttributeGroup? FirstGroup(IppGroupTag tag) {
        foreach (IppAttributeGroup group in Groups) {
            if (group.Tag == tag) return group;
        }
        return null;
    }

    public IEnumerable<IppAttributeGroup> GroupsOf(IppGroupTag tag) => Groups.Where(group => group.Tag == tag);
}
