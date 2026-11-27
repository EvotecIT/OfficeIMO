namespace OfficeIMO.Entities;

/// <summary>Local format and check validation outcomes.</summary>
public enum EntityValidationStatus {
    /// <summary>The value passed all applicable validation.</summary>
    Valid,
    /// <summary>No value was supplied.</summary>
    Empty,
    /// <summary>The value contains unsupported characters.</summary>
    InvalidCharacters,
    /// <summary>The canonical value has an unsupported length.</summary>
    InvalidLength,
    /// <summary>The value has an invalid structure.</summary>
    InvalidFormat,
    /// <summary>The defined check did not pass.</summary>
    InvalidCheckDigit,
    /// <summary>The encoded date does not exist.</summary>
    InvalidDate,
    /// <summary>The value is in a reserved or unallocated range.</summary>
    NotAllocated,
    /// <summary>The recognized variant is outside the supported contract.</summary>
    UnsupportedVariant,
}
