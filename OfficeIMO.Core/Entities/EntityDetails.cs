using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Entities;

/// <summary>Attributes directly encoded in a valid entity value.</summary>
public sealed class EntityDetails {
    /// <summary>Creates details with a read-only copy of the scheme-specific attributes.</summary>
    public EntityDetails(DateTime? birthDate = null, EntitySex? sex = null, string? country = null,
        string? variant = null, IReadOnlyDictionary<string, string>? properties = null) {
        BirthDate = birthDate;
        Sex = sex;
        Country = country;
        Variant = variant;
        var copied = new Dictionary<string, string>(StringComparer.Ordinal);
        if (properties != null) {
            foreach (KeyValuePair<string, string> property in properties) copied.Add(property.Key, property.Value);
        }
        Properties = new ReadOnlyDictionary<string, string>(copied);
    }

    /// <summary>Encoded birth date, if present.</summary>
    public DateTime? BirthDate { get; }
    /// <summary>Encoded sex, if present.</summary>
    public EntitySex? Sex { get; }
    /// <summary>ISO country code encoded in the value, if present.</summary>
    public string? Country { get; }
    /// <summary>Scheme-defined format variant, if present.</summary>
    public string? Variant { get; }
    /// <summary>Scheme-defined properties. These properties may contain sensitive data.</summary>
    public IReadOnlyDictionary<string, string> Properties { get; }
}
