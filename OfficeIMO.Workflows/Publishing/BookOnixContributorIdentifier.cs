namespace OfficeIMO.Workflows;

/// <summary>Contributor identity schemes supported by the ONIX export profile (list 44).</summary>
public enum BookOnixContributorIdentifierType {
    /// <summary>A publisher-defined scheme, requiring a distinctive scheme name (01).</summary>
    Proprietary = 1,
    /// <summary>International Standard Name Identifier, in compact ONIX form (16).</summary>
    Isni = 16,
    /// <summary>Open Researcher and Contributor ID, in compact ONIX form (21).</summary>
    Orcid = 21
}

/// <summary>
/// Explicit contributor identity. Values are retained verbatim; ISNI and ORCID require
/// 15 ASCII digits followed by a digit or uppercase X. Checksums and registration are not verified.
/// Proprietary values and scheme names are limited to 100 UTF-16 code units each.
/// </summary>
public sealed record BookOnixContributorIdentifier(BookOnixContributorIdentifierType Type, string Value, string? SchemeName = null);
