using System;

namespace OfficeIMO.Entities;

/// <summary>The local validation result for an entity value.</summary>
public sealed class EntityValidationResult {
    /// <summary>Creates a validation result. Supply a canonical value for a valid result.</summary>
    public EntityValidationResult(EntityScheme scheme, EntityValidationStatus status, string? normalized = null,
        EntityCheckOutcome checkDigit = EntityCheckOutcome.NotDefined, EntityCheckInfo? check = null) {
        Scheme = scheme ?? throw new ArgumentNullException(nameof(scheme));
        if (status == EntityValidationStatus.Valid && normalized == null) throw new ArgumentNullException(nameof(normalized));
        Status = status;
        Normalized = normalized;
        CheckDigit = checkDigit;
        Check = check;
    }

    /// <summary>Whether format and applicable checks passed.</summary>
    public bool IsValid => Status == EntityValidationStatus.Valid;
    /// <summary>Outcome of validation.</summary>
    public EntityValidationStatus Status { get; }
    /// <summary>Canonical value when the format parsed. This property may contain sensitive data.</summary>
    public string? Normalized { get; }
    /// <summary>Scheme used to validate the value.</summary>
    public EntityScheme Scheme { get; }
    /// <summary>Outcome of the checksum, when applied.</summary>
    public EntityCheckOutcome CheckDigit { get; }
    /// <summary>The check applied by the scheme, if any.</summary>
    public EntityCheckInfo? Check { get; }
    /// <summary>Provenance of the checksum applied to this variant.</summary>
    public EntityCheckProvenance CheckProvenance => Check?.Provenance ?? EntityCheckProvenance.None;

    /// <inheritdoc />
    public override string ToString() => $"{Scheme.Id}: {Status}";
}
