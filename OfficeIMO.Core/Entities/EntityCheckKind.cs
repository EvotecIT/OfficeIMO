namespace OfficeIMO.Entities;

/// <summary>Families of check digit algorithms.</summary>
public enum EntityCheckKind {
    /// <summary>No check digit.</summary>
    None,
    /// <summary>Weighted modulus 10.</summary>
    Mod10,
    /// <summary>Weighted modulus 11.</summary>
    Mod11,
    /// <summary>Modulus 23 with a letter table.</summary>
    Mod23,
    /// <summary>ISO 7064 MOD 97-10.</summary>
    Mod97_10,
    /// <summary>Luhn double-add-double.</summary>
    Luhn,
    /// <summary>Another check or a choice of checks by variant.</summary>
    Other,
}
