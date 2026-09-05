using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Immutable OpenType feature selections used by the shared text-shaping contract.
/// </summary>
/// <remarks>
/// Tags use the four-character OpenType representation. A value of zero disables a feature;
/// positive values enable it or select an alternate. Features absent from the collection retain
/// the shaping engine's script defaults.
/// </remarks>
public sealed class OfficeTextFeatureSettings : IEquatable<OfficeTextFeatureSettings> {
    private const int MaximumFeatureCount = 128;
    private readonly IReadOnlyDictionary<string, int> _features;

    /// <summary>No explicit feature overrides.</summary>
    public static OfficeTextFeatureSettings Default { get; } = new OfficeTextFeatureSettings();

    /// <summary>Creates an empty feature selection.</summary>
    public OfficeTextFeatureSettings()
        : this((IEnumerable<KeyValuePair<string, int>>?)null) {
    }

    /// <summary>Creates an immutable snapshot of feature tag values.</summary>
    public OfficeTextFeatureSettings(IEnumerable<KeyValuePair<string, int>>? features) {
        var snapshot = new Dictionary<string, int>(StringComparer.Ordinal);
        if (features != null) {
            foreach (KeyValuePair<string, int> feature in features) {
                ValidateTag(feature.Key, nameof(features));
                if (feature.Value < 0 || feature.Value > ushort.MaxValue) {
                    throw new ArgumentOutOfRangeException(nameof(features), "OpenType feature values must be between 0 and 65535.");
                }
                snapshot[feature.Key] = feature.Value;
                if (snapshot.Count > MaximumFeatureCount) {
                    throw new ArgumentException("Text shaping supports at most 128 explicit OpenType features.", nameof(features));
                }
            }
        }
        _features = new ReadOnlyDictionary<string, int>(snapshot);
    }

    /// <summary>Explicit OpenType feature values keyed by four-character tag.</summary>
    public IReadOnlyDictionary<string, int> Features => _features;

    /// <summary>True when no feature differs from the shaping engine's defaults.</summary>
    public bool IsDefault => _features.Count == 0;

    /// <summary>Returns the explicit value for a feature tag.</summary>
    public bool TryGetValue(string tag, out int value) => _features.TryGetValue(tag, out value);

    /// <summary>Returns a new immutable selection with one feature value added or replaced.</summary>
    public OfficeTextFeatureSettings With(string tag, int value) {
        ValidateTag(tag, nameof(tag));
        if (value < 0 || value > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
        var updated = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, int> feature in _features) updated[feature.Key] = feature.Value;
        updated[tag] = value;
        return new OfficeTextFeatureSettings(updated);
    }

    /// <inheritdoc />
    public bool Equals(OfficeTextFeatureSettings? other) {
        if (ReferenceEquals(this, other)) return true;
        if (other == null || _features.Count != other._features.Count) return false;
        foreach (KeyValuePair<string, int> feature in _features) {
            if (!other._features.TryGetValue(feature.Key, out int value) || value != feature.Value) return false;
        }
        return true;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as OfficeTextFeatureSettings);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            var tags = new List<string>(_features.Keys);
            tags.Sort(StringComparer.Ordinal);
            foreach (string tag in tags) {
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(tag);
                hash = hash * 31 + _features[tag];
            }
            return hash;
        }
    }

    private static void ValidateTag(string? tag, string parameterName) {
        if (tag == null || tag.Length != 4) {
            throw new ArgumentException("OpenType feature tags must contain exactly four characters.", parameterName);
        }
        for (int index = 0; index < tag.Length; index++) {
            if (tag[index] < 0x20 || tag[index] > 0x7E) {
                throw new ArgumentException("OpenType feature tags must contain printable ASCII characters.", parameterName);
            }
        }
    }
}
