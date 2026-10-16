using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Installed alternatives for missing Office document families; never grants host-font access.</summary>
internal static class OfficeDocumentFontSubstitutions {
    /// <summary>Ordered substitutes, after the requested installed family has been tried.</summary>
    internal static IEnumerable<string> Enumerate(string family) {
        switch (Normalize(family)) {
            case "calibri":
            case "calibrilight":
                yield return "Carlito";
                yield return "Arial";
                yield return "Liberation Sans";
                yield return "DejaVu Sans";
                yield break;
            case "cambria":
                yield return "Caladea";
                yield return "Times New Roman";
                yield return "Liberation Serif";
                yield return "DejaVu Serif";
                yield break;
            case "arial":
            case "helvetica":
                yield return "Liberation Sans";
                yield return "DejaVu Sans";
                yield return "Noto Sans";
                yield break;
            case "times":
            case "timesroman":
            case "timesnewroman":
                yield return "Liberation Serif";
                yield return "DejaVu Serif";
                yield return "Noto Serif";
                yield break;
            case "courier":
            case "couriernew":
                yield return "Liberation Mono";
                yield return "Menlo";
                yield return "DejaVu Sans Mono";
                yield break;
            case "aptos":
            case "aptosdisplay":
            case "aptosbody":
                yield return "Inter";
                yield return "Carlito";
                yield return "Arial";
                yield return "Liberation Sans";
                yield return "DejaVu Sans";
                yield break;
        }
    }

    /// <summary>
    /// Pairs selected with compatibility intent. This policy does not qualify the installed font
    /// versions, vertical metrics, wrapping, or pagination; other pairs remain layout-sensitive.
    /// </summary>
    internal static bool IsIntendedCompatibleSubstitute(string source, string target) {
        string normalized = Normalize(source);
        return (normalized == "calibri" && target.Equals("Carlito", StringComparison.OrdinalIgnoreCase)) ||
            (normalized == "cambria" && target.Equals("Caladea", StringComparison.OrdinalIgnoreCase)) ||
            ((normalized == "arial" || normalized == "helvetica") && target.Equals("Liberation Sans", StringComparison.OrdinalIgnoreCase)) ||
            ((normalized == "times" || normalized == "timesroman" || normalized == "timesnewroman") && target.Equals("Liberation Serif", StringComparison.OrdinalIgnoreCase)) ||
            ((normalized == "courier" || normalized == "couriernew") && target.Equals("Liberation Mono", StringComparison.OrdinalIgnoreCase));
    }

    private static string Normalize(string family) =>
        family.Trim().Replace(" ", string.Empty).Replace("-", string.Empty).ToLowerInvariant();
}
