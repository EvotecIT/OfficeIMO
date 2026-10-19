using System;

namespace OfficeIMO.Drawing;

/// <summary>Bounds the solid-paint, positioned Latin-text drawing interchange profile.</summary>
/// <remarks>The producer and consumer must register the same licensed static TrueType faces.
/// Font bytes and fallback font registries are not carried by the drawing payload.</remarks>
public sealed class OfficeDrawingPortableOptions {
    /// <summary>Creates an export profile bound to one explicit font family.</summary>
    public OfficeDrawingPortableOptions(string fontFamily) {
        if (string.IsNullOrWhiteSpace(fontFamily)) throw new ArgumentException("An explicit font family is required.", nameof(fontFamily));
        FontFamily = fontFamily.Trim();
    }

    /// <summary>Exact family name accepted for every text element, ignoring case.</summary>
    public string FontFamily { get; }
    /// <summary>Maximum source and generated drawing elements, including clip groups.</summary>
    public int MaximumElements { get; set; } = 10000;
    /// <summary>Maximum path commands across the complete drawing.</summary>
    public int MaximumCommands { get; set; } = 100000;
    /// <summary>Maximum nested drawing/group depth.</summary>
    public int MaximumDepth { get; set; } = 32;
    /// <summary>Maximum UTF-16 characters across all text runs.</summary>
    public int MaximumTextCharacters { get; set; } = 100000;
    /// <summary>Maximum characters in the returned JSON document.</summary>
    public int MaximumJsonCharacters { get; set; } = 8 * 1024 * 1024;

    internal void Validate() {
        if (MaximumElements < 1 || MaximumElements > 100000) throw new ArgumentOutOfRangeException(nameof(MaximumElements));
        if (MaximumCommands < 1 || MaximumCommands > 1000000) throw new ArgumentOutOfRangeException(nameof(MaximumCommands));
        if (MaximumDepth < 1 || MaximumDepth > 32) throw new ArgumentOutOfRangeException(nameof(MaximumDepth));
        if (MaximumTextCharacters < 1 || MaximumTextCharacters > 1000000) throw new ArgumentOutOfRangeException(nameof(MaximumTextCharacters));
        if (MaximumJsonCharacters < 1 || MaximumJsonCharacters > 32 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaximumJsonCharacters));
    }
}
