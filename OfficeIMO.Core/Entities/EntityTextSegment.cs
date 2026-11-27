using System;

namespace OfficeIMO.Entities;

/// <summary>One independent unit of text with optional location and field context.</summary>
public readonly struct EntityTextSegment {
    /// <summary>Constructs an input segment. Set structured source for a known typed value.</summary>
    public EntityTextSegment(string text, EntitySegmentLocation? location = null, string? label = null, bool structuredSource = false) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Location = location;
        Label = label;
        StructuredSource = structuredSource;
    }

    /// <summary>Text to search.</summary>
    public string Text { get; }
    /// <summary>Coordinates supplied by the host.</summary>
    public EntitySegmentLocation? Location { get; }
    /// <summary>Optional field or column label.</summary>
    public string? Label { get; }
    /// <summary>Whether the whole segment represents a typed field value.</summary>
    public bool StructuredSource { get; }
}
