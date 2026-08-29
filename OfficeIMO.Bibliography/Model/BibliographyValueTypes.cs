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
    /// <summary>Initializes an identifier.</summary>
    public BibliographyIdentifier(string scheme, string value) {
        Scheme = string.IsNullOrWhiteSpace(scheme) ? throw new ArgumentException("Identifier scheme cannot be empty.", nameof(scheme)) : scheme.Trim();
        Value = string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("Identifier value cannot be empty.", nameof(value)) : value.Trim();
    }

    /// <summary>Identifier scheme.</summary>
    public string Scheme { get; set; }
    /// <summary>Identifier value.</summary>
    public string Value { get; set; }
}

/// <summary>An ordered native field retained outside the typed model.</summary>
public sealed class BibliographyNativeField {
    /// <summary>Initializes a native field.</summary>
    public BibliographyNativeField(BibliographyFormat format, string name, string value, string? rawValue = null) {
        Format = format;
        Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentException("Field name cannot be empty.", nameof(name)) : name;
        Value = value ?? throw new ArgumentNullException(nameof(value));
        RawValue = rawValue;
    }

    /// <summary>Source format that owns the field name and syntax.</summary>
    public BibliographyFormat Format { get; }
    /// <summary>Native field, tag, or element name.</summary>
    public string Name { get; }
    /// <summary>Decoded semantic value.</summary>
    public string Value { get; set; }
    /// <summary>Optional raw source representation.</summary>
    public string? RawValue { get; }
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
