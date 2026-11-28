namespace OfficeIMO.Entities;

/// <summary>Publication status of a check applied by a scheme.</summary>
public enum EntityCheckProvenance {
    /// <summary>No checksum provenance applies.</summary>
    None,
    /// <summary>Published by an issuer, standards body, or government authority.</summary>
    IssuerPublished,
    /// <summary>Documented independently of the issuing authority.</summary>
    CommunityDocumented,
}
