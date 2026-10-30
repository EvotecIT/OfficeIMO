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

    /// <summary>Returns the generic CSS family (<c>sans-serif</c>, <c>serif</c> or <c>monospace</c>) that best describes a family name.</summary>
    internal static string ClassifyGenericFamily(string family) {
        string key = Normalize(family);
        if (IsMonospace(key)) return "monospace";
        if (IsSerif(key)) return "serif";
        return "sans-serif";
    }

    private static bool IsMonospace(string key) =>
        key.Contains("mono") || key.Contains("courier") || key.Contains("consol") || key.Contains("code") ||
        key == "menlo" || key == "monaco" || key == "sfmono";

    private static bool IsSerif(string key) =>
        !key.Contains("sans") && (key.Contains("serif") || key.Contains("times") || key.Contains("roman") ||
            key.Contains("georgia") || key.Contains("cambria") || key.Contains("caladea") || key.Contains("gelasio") ||
            key.Contains("garamond") || key.Contains("baskerville") || key.Contains("palatino") || key.Contains("bookantiqua") ||
            key.Contains("century") || key.Contains("tinos") || key.Contains("constantia") || key.Contains("didot") ||
            key.Contains("minion") || key.Contains("charter"));

    private static string Normalize(string family) =>
        family.Trim().Replace(" ", string.Empty).Replace("-", string.Empty).ToLowerInvariant();
}
