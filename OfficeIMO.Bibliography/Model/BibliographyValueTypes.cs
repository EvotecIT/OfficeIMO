namespace OfficeIMO.Bibliography;

/// <summary>A personal or literal organization name.</summary>
public sealed class BibliographyName {
    /// <summary>Given name.</summary>
    public string? Given { get; set; }
    /// <summary>Family name.</summary>
    public string? Family { get; set; }
    /// <summary>Literal corporate or otherwise indivisible name.</summary>
    public string? Literal { get; set; }
    /// <summary>Name suffix.</summary>
    public string? Suffix { get; set; }
    /// <summary>Dropping particle.</summary>
    public string? DroppingParticle { get; set; }
    /// <summary>Non-dropping particle.</summary>
    public string? NonDroppingParticle { get; set; }
    /// <summary>Unknown native name properties in source order.</summary>
    public IList<BibliographyNativeField> NativeFields { get; } = new List<BibliographyNativeField>();

    /// <summary>Returns a stable display representation.</summary>
    public override string ToString() {
        if (!string.IsNullOrWhiteSpace(Literal)) return Literal!;
        string family = JoinNonEmpty(" ", NonDroppingParticle, Family);
        string given = JoinNonEmpty(" ", Given, DroppingParticle);
        string name = JoinNonEmpty(", ", family, given);
        return JoinNonEmpty(", ", name, Suffix);
    }

    private static string JoinNonEmpty(string separator, params string?[] values) =>
        string.Join(separator, values.Where(static value => !string.IsNullOrWhiteSpace(value)).Select(static value => value!.Trim()));
}

/// <summary>A contributor and their role.</summary>
public sealed class BibliographyContributor {
    /// <summary>Initializes a contributor.</summary>
    public BibliographyContributor(BibliographyContributorRole role, BibliographyName name) {
        Role = role;
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>Contributor role.</summary>
    public BibliographyContributorRole Role { get; set; }
    /// <summary>Contributor name.</summary>
    public BibliographyName Name { get; }
}

/// <summary>A possibly partial or literal bibliographic date.</summary>
public sealed class BibliographyDate {
    /// <summary>Date role.</summary>
    public BibliographyDateRole Role { get; set; } = BibliographyDateRole.Issued;
    /// <summary>Four-digit or source-provided year.</summary>
    public int? Year { get; set; }
    /// <summary>Month number from 1 through 12.</summary>
    public int? Month { get; set; }
    /// <summary>Day number from 1 through 31.</summary>
    public int? Day { get; set; }
    /// <summary>End year for a date range.</summary>
    public int? EndYear { get; set; }
    /// <summary>End month number from 1 through 12 for a date range.</summary>
    public int? EndMonth { get; set; }
    /// <summary>End day number from 1 through 31 for a date range.</summary>
    public int? EndDay { get; set; }
    /// <summary>Literal date when numeric parts are insufficient.</summary>
    public string? Literal { get; set; }
    /// <summary>Unknown native date properties in source order.</summary>
    public IList<BibliographyNativeField> NativeFields { get; } = new List<BibliographyNativeField>();
}

/// <summary>A typed identifier such as DOI, PMID, ISBN, or ISSN.</summary>
public sealed class BibliographyIdentifier {
    private string _scheme;
    private string _value;

    /// <summary>Initializes an identifier.</summary>
    public BibliographyIdentifier(string scheme, string value) {
        _scheme = ValidateScheme(scheme, nameof(scheme));
        _value = ValidateValue(value, nameof(value));
    }

    /// <summary>Identifier scheme.</summary>
    public string Scheme { get => _scheme; set => _scheme = ValidateScheme(value, nameof(value)); }
    /// <summary>Identifier value, preserving nonempty source whitespace.</summary>
    public string Value { get => _value; set => _value = ValidateValue(value, nameof(value)); }

    private static string ValidateScheme(string? value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Identifier scheme cannot be empty.", parameterName);
        if (value!.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0) throw new ArgumentException("Identifier scheme cannot contain line breaks.", parameterName);
        return value.Trim();
    }

    private static string ValidateValue(string? value, string parameterName) => string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("Identifier value cannot be empty.", parameterName) : value!;
}

/// <summary>An ordered native field retained outside the typed model.</summary>
public sealed class BibliographyNativeField {
    private readonly string _originalValue;
    private readonly bool _rawValueRepresentsOriginalValue;

    /// <summary>Initializes a native field.</summary>
    public BibliographyNativeField(BibliographyFormat format, string name, string value, string? rawValue = null)
        : this(format, name, value, rawValue, false) { }

    private BibliographyNativeField(BibliographyFormat format, string name, string value, string? rawValue, bool allowEmptyName) {
        Format = format;
        Name = !allowEmptyName && string.IsNullOrWhiteSpace(name) ? throw new ArgumentException("Field name cannot be empty.", nameof(name)) : name ?? throw new ArgumentNullException(nameof(name));
        Value = value ?? throw new ArgumentNullException(nameof(value));
        _originalValue = value;
        RawValue = rawValue;
        _rawValueRepresentsOriginalValue = RawValueRepresentsValue(format, Name, value, rawValue);
    }

    internal static BibliographyNativeField FromParsedSource(BibliographyFormat format, string name, string value, string? rawValue = null) =>
        new BibliographyNativeField(format, name, value, rawValue, true);

    /// <summary>Source format that owns the field name and syntax.</summary>
    public BibliographyFormat Format { get; }
    /// <summary>Native field, tag, or element name.</summary>
    public string Name { get; }
    /// <summary>Decoded semantic value.</summary>
    public string Value { get; set; }
    /// <summary>Optional raw source representation.</summary>
    public string? RawValue { get; }

    internal string? UnmodifiedRawValue => RawValue != null && _rawValueRepresentsOriginalValue && string.Equals(Value, _originalValue, StringComparison.Ordinal) ? RawValue : null;
    internal bool HasInconsistentRawValue => RawValue != null && !_rawValueRepresentsOriginalValue && string.Equals(Value, _originalValue, StringComparison.Ordinal);

    private static bool RawValueRepresentsValue(BibliographyFormat format, string name, string value, string? rawValue) {
        if (rawValue == null) return true;
        try {
            if (format == BibliographyFormat.CslJson) {
                using System.Text.Json.JsonDocument document = System.Text.Json.JsonDocument.Parse(rawValue, new System.Text.Json.JsonDocumentOptions { MaxDepth = CslJsonCodec.NativeJsonMaximumDepth });
                System.Text.Json.JsonElement root = document.RootElement;
                switch (root.ValueKind) {
                    case System.Text.Json.JsonValueKind.String: return string.Equals(root.GetString() ?? string.Empty, value, StringComparison.Ordinal);
                    case System.Text.Json.JsonValueKind.Number:
                    case System.Text.Json.JsonValueKind.True:
                    case System.Text.Json.JsonValueKind.False: return string.Equals(root.GetRawText(), value, StringComparison.Ordinal);
                    case System.Text.Json.JsonValueKind.Null:
                    case System.Text.Json.JsonValueKind.Undefined: return value.Length == 0;
                    default: return string.Equals(root.GetRawText(), value, StringComparison.Ordinal);
                }
            }
            if (format == BibliographyFormat.EndNoteXml) {
                System.Xml.Linq.XElement element = System.Xml.Linq.XElement.Parse(rawValue, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                return string.Equals(element.Name.LocalName, name, StringComparison.Ordinal) && string.Equals(element.Value, value, StringComparison.Ordinal);
            }
            return true;
        } catch (Exception exception) when (exception is System.Text.Json.JsonException || exception is System.Xml.XmlException || exception is InvalidOperationException || exception is ArgumentException) {
            return false;
        }
    }
}

/// <summary>An ordered document-level directive or value outside citation records.</summary>
public sealed class BibliographyNativeEntry {
    /// <summary>Initializes a document-level native entry.</summary>
    public BibliographyNativeEntry(BibliographyFormat format, string kind, string value, string? name = null) {
        Format = format;
        Kind = string.IsNullOrWhiteSpace(kind) ? throw new ArgumentException("Entry kind cannot be empty.", nameof(kind)) : kind;
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Name = name;
    }

    /// <summary>Owning source format.</summary>
    public BibliographyFormat Format { get; }
    /// <summary>Native directive or entry kind.</summary>
    public string Kind { get; }
    /// <summary>Optional directive name.</summary>
    public string? Name { get; }
    /// <summary>Decoded or retained value.</summary>
    public string Value { get; set; }
}
