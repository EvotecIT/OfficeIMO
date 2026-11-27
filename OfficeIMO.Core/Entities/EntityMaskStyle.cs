namespace OfficeIMO.Entities;

/// <summary>Replacement style for a selected finding.</summary>
public enum EntityMaskStyle {
    /// <summary>Uses the entity kind as a bracketed label.</summary>
    KindLabel,
    /// <summary>Uses a configurable repeated character.</summary>
    Fixed,
    /// <summary>Preserves the portion explicitly allowed by the scheme.</summary>
    Partial,
    /// <summary>Uses a deterministic keyed token.</summary>
    Pseudonym,
}
