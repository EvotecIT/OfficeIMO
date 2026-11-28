using System;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Defines how to validate and present one family of entity values.</summary>
public abstract class EntityScheme {
    private IReadOnlyList<EntityCheckInfo>? _checks;

    /// <summary>Stable catalog identifier.</summary>
    public abstract string Id { get; }
    /// <summary>Category used for filtering and masking.</summary>
    public abstract EntityKind Kind { get; }
    /// <summary>Human-readable scheme name.</summary>
    public abstract string DisplayName { get; }
    /// <summary>ISO 3166-1 alpha-2 country, or null for an international scheme.</summary>
    public virtual string? Country => null;
    /// <summary>Organization responsible for issuing or standardizing these values.</summary>
    public virtual string? Issuer => null;
    /// <summary>Documentation for the implemented format or check.</summary>
    public virtual Uri? Specification => null;
    /// <summary>Minimum normalized length in UTF-16 code units.</summary>
    public virtual int MinLength => 1;
    /// <summary>Maximum normalized length in UTF-16 code units.</summary>
    public virtual int MaxLength => 64;
    /// <summary>Check digit algorithm, if defined.</summary>
    public virtual EntityCheckKind Check => EntityCheckKind.None;
    /// <summary>Checks available for the supported formats and variants.</summary>
    public virtual IReadOnlyList<EntityCheckInfo> Checks {
        get {
            IReadOnlyList<EntityCheckInfo>? existing = Volatile.Read(ref _checks);
            if (existing != null) return existing;
            IReadOnlyList<EntityCheckInfo> checks = DescribeChecks();
            return Interlocked.CompareExchange(ref _checks, checks, null) ?? checks;
        }
    }

    /// <summary>Describes the source status used by the default single checksum.</summary>
    protected virtual EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.IssuerPublished;

    /// <summary>Describes the known checks. Override when formats use different algorithms or sources.</summary>
    protected virtual IReadOnlyList<EntityCheckInfo> DescribeChecks() => Check == EntityCheckKind.None
        ? Array.Empty<EntityCheckInfo>()
        : Array.AsReadOnly(new[] { new EntityCheckInfo(DisplayName + " checksum", Check, DefaultCheckProvenance, Specification) });

    /// <summary>Default masking classification.</summary>
    public virtual EntitySensitivity Sensitivity => EntitySensitivity.Personal;
    /// <summary>Scheme identifiers that can appear inside this scheme's match.</summary>
    public virtual IReadOnlyList<string> Contains => Array.Empty<string>();

    internal virtual void RecognizeContained(EntityRecognitionContext context, ScoredEntity reading) { }

    /// <summary>Checks written or normalized input without throwing for a malformed value.</summary>
    public abstract EntityValidationResult Validate(string? value);

    /// <summary>Checks whether a value passes local validation.</summary>
    public bool IsValid(string? value) => Validate(value).IsValid;

    /// <summary>Returns the canonical value on successful validation.</summary>
    public bool TryNormalize(string? value, out string normalized) {
        EntityValidationResult result = Validate(value);
        normalized = result.IsValid ? result.Normalized! : string.Empty;
        return result.IsValid;
    }

    /// <summary>Returns the canonical value or null if validation fails.</summary>
    public string? Normalize(string? value) => TryNormalize(value, out string normalized) ? normalized : null;

    /// <summary>Formats a canonical value for the requested use.</summary>
    public virtual string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) =>
        normalized ?? throw new ArgumentNullException(nameof(normalized));

    /// <summary>Returns the attributes encoded in a valid value, when supported.</summary>
    public virtual EntityDetails? Decode(string normalized) => null;

    /// <summary>Masks a canonical value, preserving only the scheme's permitted display fragments.</summary>
    public virtual string MaskPartial(string normalized) =>
        new('•', normalized?.Length ?? throw new ArgumentNullException(nameof(normalized)));
}
