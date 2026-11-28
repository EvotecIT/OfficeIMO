using System.Collections.Generic;

namespace OfficeIMO.Entities;

/// <summary>Settings for a single masking operation.</summary>
public sealed class EntityMaskOptions {
    /// <summary>Replacement style.</summary>
    public EntityMaskStyle Style { get; set; } = EntityMaskStyle.KindLabel;
    /// <summary>Character repeated for fixed masking.</summary>
    public char FixedCharacter { get; set; } = '█';
    /// <summary>Fixed replacement length for string input, or null to preserve the written length.</summary>
    public int? FixedLength { get; set; }
    /// <summary>Secret HMAC key required for keyed pseudonyms.</summary>
    public byte[]? PseudonymKey { get; set; }
    /// <summary>Sensitivities selected for masking.</summary>
    public ISet<EntitySensitivity> Sensitivities { get; } = new HashSet<EntitySensitivity> {
        EntitySensitivity.Personal, EntitySensitivity.Financial,
    };
}
