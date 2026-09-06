using System;

namespace OfficeIMO.Drawing;

public readonly partial struct OfficeColor {
    // Computed CSS commonly serializes opaque colors as rgba(r, g, b, 1).
    // Read integer forms without temporary strings; other syntax uses the full parser.
    private static bool TryParseIntegerRgb(string value, out OfficeColor color) {
        color = default;
        bool alpha = value.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase);
        if (!alpha && !value.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase)) return false;
        int index = alpha ? 5 : 4;
        if (!TryReadIntegerRgbChannel(value, ref index, ',', out byte red)
            || !TryReadIntegerRgbChannel(value, ref index, ',', out byte green)
            || !TryReadIntegerRgbChannel(value, ref index, alpha ? ',' : ')', out byte blue)) return false;
        byte opacity = 1;
        if (alpha && (!TryReadIntegerRgbChannel(value, ref index, ')', out opacity) || opacity > 1)) return false;
        if (index != value.Length) return false;
        color = FromRgba(red, green, blue, opacity == 1 ? (byte)255 : (byte)0);
        return true;
    }

    private static bool TryReadIntegerRgbChannel(string value, ref int index, char separator, out byte channel) {
        channel = 0;
        while (index < value.Length && char.IsWhiteSpace(value[index])) index++;
        int start = index;
        int number = 0;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9') {
            number = number * 10 + value[index++] - '0';
            if (number > 255) return false;
        }
        if (index == start) return false;
        while (index < value.Length && char.IsWhiteSpace(value[index])) index++;
        if (index >= value.Length || value[index++] != separator) return false;
        channel = (byte)number;
        return true;
    }
}
