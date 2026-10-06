using System;

namespace OfficeIMO.Drawing;

/// <summary>Unicode mathematical token presentation shared by format adapters.</summary>
internal static class OfficeMathTextTransform {
    // MathML Core section 4.2 and appendix C.1. This is a scalar mapping, not a
    // font-style request, normalization or a transformation of multi-letter names.
    internal static string MathAuto(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (text.Length != 1) return text;
        int scalar = text[0];
        int mapped;
        if (scalar >= 'A' && scalar <= 'Z') mapped = scalar + 0x1D3F3;
        else if (scalar >= 'a' && scalar <= 'z') mapped = scalar == 'h' ? 0x210E : scalar + 0x1D3ED;
        else if (scalar >= 0x0391 && scalar <= 0x03A9 && scalar != 0x03A2) mapped = scalar + 0x1D351;
        else if (scalar >= 0x03B1 && scalar <= 0x03C9) mapped = scalar + 0x1D34B;
        else mapped = scalar switch {
            0x0131 => 0x1D6A4,
            0x0237 => 0x1D6A5,
            0x03F4 => 0x1D6F3,
            0x2207 => 0x1D6FB,
            0x2202 => 0x1D715,
            0x03F5 => 0x1D716,
            0x03D1 => 0x1D717,
            0x03F0 => 0x1D718,
            0x03D5 => 0x1D719,
            0x03F1 => 0x1D71A,
            0x03D6 => 0x1D71B,
            _ => scalar
        };
        return mapped == scalar ? text : char.ConvertFromUtf32(mapped);
    }
}
