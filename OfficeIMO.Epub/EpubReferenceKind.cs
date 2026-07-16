namespace OfficeIMO.Epub;

/// <summary>Identifies how an EPUB URL reference resolves.</summary>
public enum EpubReferenceKind {
    /// <summary>The reference could not be resolved safely.</summary>
    Invalid,

    /// <summary>The reference resolves to a case-sensitive path in the EPUB container.</summary>
    Container,

    /// <summary>The reference resolves outside the EPUB container.</summary>
    External,

    /// <summary>The reference is an embedded <c>data:</c> URL.</summary>
    Data
}
