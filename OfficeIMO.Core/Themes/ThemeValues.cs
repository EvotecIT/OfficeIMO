using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

internal static class ThemeValues {
    internal static string FontName(string value, string name) => string.IsNullOrWhiteSpace(value)
        ? throw new ArgumentException("A name is required.", name) : value;
    internal static OfficeLength NonNegative(OfficeLength value, string name) => value < OfficeLength.Zero
        ? throw new ArgumentOutOfRangeException(name, "Theme sizes and spacing cannot be negative.") : value;
    internal static void Literal(OfficeColor value, string name) {
        if (value.IsThemeRole) throw new ArgumentException("Palette colors must be literal to prevent circular role resolution.", name);
    }
    internal static IReadOnlyList<string> Names(IReadOnlyList<string> values, string name) {
        if (values == null) throw new ArgumentNullException(name);
        var copy = new List<string>(values.Count);
        foreach (string value in values) copy.Add(FontName(value, name));
        return copy.AsReadOnly();
    }
    internal static IReadOnlyList<OfficeColor> Colors(IReadOnlyList<OfficeColor> values, string name) {
        if (values == null) throw new ArgumentNullException(name);
        if (values.Count == 0) throw new ArgumentException("A series palette must contain at least one color.", name);
        return new List<OfficeColor>(values).AsReadOnly();
    }
}
