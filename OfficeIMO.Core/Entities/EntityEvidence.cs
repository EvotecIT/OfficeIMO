using System;

namespace OfficeIMO.Entities;

/// <summary>Evidence considered when scoring a reading.</summary>
[Flags]
public enum EntityEvidence {
    /// <summary>No evidence flags.</summary>
    None = 0,
    /// <summary>The scheme's checksum passed.</summary>
    CheckDigitPassed = 1,
    /// <summary>This variant has no applied checksum.</summary>
    NoCheckDefined = 1 << 1,
    /// <summary>The value follows the printed issuer grouping.</summary>
    IssuerFormatting = 1 << 2,
    /// <summary>The input includes the scheme's country prefix.</summary>
    CountryPrefix = 1 << 3,
    /// <summary>A keyword in the same segment supports the scheme.</summary>
    ContextKeyword = 1 << 4,
    /// <summary>The host-supplied segment label supports the scheme.</summary>
    Label = 1 << 5,
    /// <summary>The decoded birth date is within the supported plausibility window.</summary>
    PlausibleDate = 1 << 6,
    /// <summary>The value starts in a recognized issuer range.</summary>
    KnownIssuerRange = 1 << 7,
    /// <summary>Another scheme also validates this span.</summary>
    AmbiguousReading = 1 << 8,
    /// <summary>The host identified this segment as a typed source field.</summary>
    StructuredSource = 1 << 9,
    /// <summary>The passed check is documented independently of the issuer.</summary>
    CommunityDocumentedCheck = 1 << 10,
}
