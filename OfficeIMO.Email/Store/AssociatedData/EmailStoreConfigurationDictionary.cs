using System.Xml.Linq;

namespace OfficeIMO.Email.Store;

/// <summary>Supported typed value in an Outlook roaming configuration dictionary.</summary>
public enum EmailStoreConfigurationValueKind {
    /// <summary>An unsupported or malformed type code retained as raw text.</summary>
    Unknown = 0,
    /// <summary>Boolean type code 3.</summary>
    Boolean = 3,
    /// <summary>signed Int32 type code 9.</summary>
    Integer = 9,
    /// <summary>Unicode string type code 18.</summary>
    String = 18
}
/// <summary>One typed roaming dictionary value with its original encoding.</summary>
public sealed class EmailStoreConfigurationValue {
    internal EmailStoreConfigurationValue(EmailStoreConfigurationValueKind kind, object? value,
        int typeCode, string raw) {
        Kind = kind;
        Value = value;
        TypeCode = typeCode;
        Raw = raw;
    }

    /// <summary>Decoded supported kind.</summary>
    public EmailStoreConfigurationValueKind Kind { get; }
    /// <summary>Decoded bool, Int32, or string; null for unsupported/malformed values.</summary>
    public object? Value { get; }
    /// <summary>Numeric type code retained from the XML.</summary>
    public int TypeCode { get; }
    /// <summary>Exact encoded value, including its type-code prefix.</summary>
    public string Raw { get; }

    /// <summary>Attempts to return the decoded value as the requested managed type.</summary>
    public bool TryGet<T>(out T? value) {
        if (Value is T exact) {
            value = exact;
            return true;
        }
        value = default;
        return false;
    }
}

/// <summary>
/// Lossless typed editor for the documented Outlook roaming dictionary XML schema.
/// Unsupported entries and unknown XML nodes are retained.
/// </summary>
public sealed class EmailStoreConfigurationDictionary {
    private static readonly XNamespace DictionaryNamespace = "dictionary.xsd";
    private readonly XDocument _xml;

    private EmailStoreConfigurationDictionary(XDocument xml) {
        _xml = xml;
    }

    /// <summary>Creates an empty dictionary with an OfficeIMO producer version.</summary>
    public static EmailStoreConfigurationDictionary Create(string producerVersion = "OfficeIMO.1") {
        if (string.IsNullOrWhiteSpace(producerVersion)) throw new ArgumentException("A producer version is required.", nameof(producerVersion));
        var root = new XElement(DictionaryNamespace + "UserConfiguration",
            new XElement(DictionaryNamespace + "Info", new XAttribute("version", producerVersion)),
            new XElement(DictionaryNamespace + "Data"));
        return new EmailStoreConfigurationDictionary(
            new XDocument(new XDeclaration("1.0", "utf-8", null), root));
    }

    /// <summary>Parses one bounded PidTagRoamingDictionary stream.</summary>
    public static EmailStoreConfigurationDictionary Parse(byte[] bytes, int maxXmlBytes = 4 * 1024 * 1024) {
        XDocument xml = EmailStoreConfigurationXml.Parse(bytes, maxXmlBytes, "The roaming dictionary");
        if (xml.Root == null || xml.Root.Name != DictionaryNamespace + "UserConfiguration" ||
            xml.Root.Element(DictionaryNamespace + "Info") == null ||
            xml.Root.Element(DictionaryNamespace + "Data") == null) {
            throw new InvalidDataException("The roaming dictionary does not have the dictionary.xsd envelope.");
        }
        return new EmailStoreConfigurationDictionary(xml);
    }

    /// <summary>Producer/version string declared by the dictionary.</summary>
    public string? ProducerVersion => (string?)Root.Element(DictionaryNamespace + "Info")?.Attribute("version");

    /// <summary>Decoded entries keyed case-sensitively as required by the XML uniqueness contract.</summary>
    public IReadOnlyDictionary<string, EmailStoreConfigurationValue> Entries {
        get {
            var result = new Dictionary<string, EmailStoreConfigurationValue>(StringComparer.Ordinal);
            foreach (XElement element in Data.Elements(DictionaryNamespace + "e")) {
                string rawKey = (string?)element.Attribute("k") ?? string.Empty;
                string key = Decode(rawKey).Value as string ?? rawKey;
                if (!result.ContainsKey(key)) result.Add(key, Decode((string?)element.Attribute("v") ?? string.Empty));
            }
            return result;
        }
    }

    /// <summary>True when the envelope, producer version, keys, and all values use supported valid encodings.</summary>
    public bool IsProtocolEnvelopeValid {
        get {
            XElement[] entries = Data.Elements(DictionaryNamespace + "e").ToArray();
            string[] keys = entries.Select(entry => Decode((string?)entry.Attribute("k") ?? string.Empty))
                .Where(value => value.Kind == EmailStoreConfigurationValueKind.String)
                .Select(value => (string)value.Value!).ToArray();
            return !string.IsNullOrWhiteSpace(ProducerVersion) && keys.Length == entries.Length &&
                keys.Distinct(StringComparer.Ordinal).Count() == keys.Length &&
                entries.All(entry => Decode((string?)entry.Attribute("v") ?? string.Empty).Kind !=
                    EmailStoreConfigurationValueKind.Unknown);
        }
    }

    /// <summary>Gets one decoded entry.</summary>
    public bool TryGet(string key, out EmailStoreConfigurationValue? value) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return Entries.TryGetValue(key, out value);
    }

    /// <summary>Creates or replaces a Boolean entry.</summary>
    public void Set(string key, bool value) => SetEncoded(key, string.Concat("3-", value ? "True" : "False"));

    /// <summary>Creates or replaces a signed Int32 entry.</summary>
    public void Set(string key, int value) =>
        SetEncoded(key, string.Concat("9-", value.ToString(CultureInfo.InvariantCulture)));

    /// <summary>Creates or replaces a Unicode string entry.</summary>
    public void Set(string key, string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        SetEncoded(key, string.Concat("18-", value));
    }

    /// <summary>Removes one exact dictionary key.</summary>
    public bool Remove(string key) {
        ValidateKey(key);
        XElement? element = FindElement(key);
        if (element == null) return false;
        element.Remove();
        return true;
    }

    /// <summary>Serializes the current lossless dictionary XML.</summary>
    public byte[] ToXml() {
        if (!IsProtocolEnvelopeValid) {
            throw new InvalidOperationException("The roaming dictionary contains an invalid envelope, key, or value.");
        }
        return EmailStoreConfigurationXml.Serialize(new XDocument(_xml));
    }

    internal static EmailStoreConfigurationValue Decode(string raw) {
        int separator = raw.IndexOf('-');
        if (separator <= 0 || !int.TryParse(raw.Substring(0, separator), NumberStyles.Integer,
            CultureInfo.InvariantCulture, out int typeCode)) {
            return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.Unknown, null, 0, raw);
        }
        string encoded = raw.Substring(separator + 1);
        switch (typeCode) {
            case 3:
                if (string.Equals(encoded, "True", StringComparison.Ordinal)) {
                    return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.Boolean, true, typeCode, raw);
                }
                if (string.Equals(encoded, "False", StringComparison.Ordinal)) {
                    return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.Boolean, false, typeCode, raw);
                }
                break;
            case 9:
                if (int.TryParse(encoded, NumberStyles.Integer, CultureInfo.InvariantCulture, out int integer)) {
                    return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.Integer, integer, typeCode, raw);
                }
                break;
            case 18:
                return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.String, encoded, typeCode, raw);
        }
        return new EmailStoreConfigurationValue(EmailStoreConfigurationValueKind.Unknown, null, typeCode, raw);
    }

    private XElement Root => _xml.Root!;
    private XElement Data => Root.Element(DictionaryNamespace + "Data")!;

    private void SetEncoded(string key, string value) {
        ValidateKey(key);
        XElement? element = FindElement(key);
        if (element == null) {
            element = new XElement(DictionaryNamespace + "e");
            Data.Add(element);
        }
        element.SetAttributeValue("k", string.Concat("18-", key));
        element.SetAttributeValue("v", value);
    }

    private XElement? FindElement(string key) => Data.Elements(DictionaryNamespace + "e")
        .FirstOrDefault(element => string.Equals(
            Decode((string?)element.Attribute("k") ?? string.Empty).Value as string,
            key, StringComparison.Ordinal));

    private static void ValidateKey(string key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (key.Length == 0) throw new ArgumentException("A configuration key cannot be empty.", nameof(key));
    }
}
