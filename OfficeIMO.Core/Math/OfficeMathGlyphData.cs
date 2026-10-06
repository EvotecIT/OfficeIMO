using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Static glyph-specific data owned by the first-party font readers.</summary>
internal interface IOfficeMathGlyphProgram {
    OfficeMathGlyphData? MathGlyphData { get; }
}

/// <summary>Detached design-unit records. Device and variation corrections are not applied.</summary>
internal sealed class OfficeMathGlyphData {
    internal int MinimumOverlap { get; set; }
    internal Dictionary<int, OfficeMathGlyphConstruction> Vertical { get; } = new();
    internal Dictionary<int, OfficeMathGlyphConstruction> Horizontal { get; } = new();
    internal Dictionary<int, int> Italics { get; } = new();
    internal Dictionary<int, int> Accents { get; } = new();
    internal Dictionary<int, OfficeMathKernTable?[]> Kerns { get; } = new();
}

internal sealed class OfficeMathGlyphConstruction {
    internal OfficeMathGlyphConstruction(OfficeMathGlyphVariant[] variants, OfficeMathGlyphPart[] parts, int italicCorrection) {
        Variants = variants; Parts = parts; ItalicCorrection = italicCorrection;
    }
    internal OfficeMathGlyphVariant[] Variants { get; }
    internal OfficeMathGlyphPart[] Parts { get; }
    internal int ItalicCorrection { get; }
}

internal readonly struct OfficeMathGlyphVariant {
    internal OfficeMathGlyphVariant(int glyph, int advance) { Glyph = glyph; Advance = advance; }
    internal int Glyph { get; }
    internal int Advance { get; }
}

internal readonly struct OfficeMathGlyphPart {
    internal OfficeMathGlyphPart(int glyph, int start, int end, int advance, bool extender) {
        Glyph = glyph; Start = start; End = end; Advance = advance; Extender = extender;
    }
    internal int Glyph { get; }
    internal int Start { get; }
    internal int End { get; }
    internal int Advance { get; }
    internal bool Extender { get; }
}

internal sealed class OfficeMathKernTable {
    internal OfficeMathKernTable(int[] heights, int[] values) { Heights = heights; Values = values; }
    internal int[] Heights { get; }
    internal int[] Values { get; }
    internal int At(double height) {
        int low = 0, high = Heights.Length;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (height >= Heights[middle]) low = middle + 1; else high = middle;
        }
        return Values[low];
    }
}
