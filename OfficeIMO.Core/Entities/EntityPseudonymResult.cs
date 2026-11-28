using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Entities;

/// <summary>Masked text with a caller-owned map for keyed replacements.</summary>
public sealed class EntityPseudonymResult {
    internal EntityPseudonymResult(string text, Dictionary<string, string> map) {
        Text = text;
        Map = new ReadOnlyDictionary<string, string>(map);
    }

    /// <summary>Text after masking selected findings.</summary>
    public string Text { get; }
    /// <summary>Maps tokens to canonical values or the written text of a combined span. Treat this map as sensitive.</summary>
    public IReadOnlyDictionary<string, string> Map { get; }
}
