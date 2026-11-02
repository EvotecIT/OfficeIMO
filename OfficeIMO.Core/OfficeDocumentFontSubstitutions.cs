using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Installed alternatives for missing Office document families; never grants host-font access.</summary>
/// <remarks>
/// The Office families used by the built-in themes take their order from the theme fallback lists
/// (<c>Build/Visual/themes.json</c>), so PDF resolution, HTML font stacks and Word alternate names agree.
/// Metric-compatible open families map back to their Office originals, and common web and UI families map onto the
/// Modern body chain. Unknown families and generic CSS names have no substitutes here, so converters keep their own
/// fallback order; <see cref="EnumerateGeneric"/> supplies the chains for generic families and standard PDF slots.
/// </remarks>
internal static class OfficeDocumentFontSubstitutions {
    private static readonly string[] SansChain = { "Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "Arimo", "Helvetica", "DejaVu Sans", "Noto Sans" };
    private static readonly string[] SerifChain = { "Cambria", "Caladea", "Georgia", "Liberation Serif", "Times New Roman", "Tinos", "DejaVu Serif", "Noto Serif" };
    private static readonly string[] MonospaceChain = { "Consolas", "Menlo", "Courier New", "Liberation Mono", "Cousine", "DejaVu Sans Mono", "Noto Sans Mono" };

    private static readonly string[] StandardSansChain = { "Arial", "Liberation Sans", "Arimo", "Helvetica", "Helvetica Neue", "Aptos", "Calibri", "Carlito", "DejaVu Sans", "Noto Sans" };
    private static readonly string[] StandardSerifChain = { "Times New Roman", "Liberation Serif", "Tinos", "Times", "Cambria", "Caladea", "Georgia", "DejaVu Serif", "Noto Serif" };
    private static readonly string[] StandardMonospaceChain = { "Courier New", "Liberation Mono", "Cousine", "Courier", "Consolas", "Menlo", "DejaVu Sans Mono", "Noto Sans Mono" };

    private static readonly Dictionary<string, string[]> Chains = BuildChains();

    /// <summary>Ordered substitutes, after the requested installed family has been tried.</summary>
    internal static IEnumerable<string> Enumerate(string family) {
        string key = Normalize(family);
        if (key.Length == 0 || IsSymbolFamily(key)) yield break;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { family.Trim() };
        if (Chains.TryGetValue(key, out string[]? chain)) {
            foreach (string candidate in chain) {
                if (seen.Add(candidate)) yield return candidate;
            }
        }
    }

    /// <summary>
    /// Ordered installed families for a generic family (<c>sans-serif</c>, <c>serif</c>, <c>monospace</c>) or a standard
    /// PDF family (<c>Helvetica</c>, <c>Times</c>, <c>Courier</c>). Standard families list their metric-compatible faces first.
    /// </summary>
    internal static IEnumerable<string> EnumerateGeneric(string family) {
        string key = Normalize(family);
        string[] chain = key switch {
            "helvetica" => StandardSansChain,
            "times" or "timesroman" => StandardSerifChain,
            "courier" => StandardMonospaceChain,
            _ => IsMonospace(key) ? MonospaceChain : IsSerif(key) ? SerifChain : SansChain
        };
        return chain;
    }

    /// <summary>Returns the generic CSS family (<c>sans-serif</c>, <c>serif</c> or <c>monospace</c>) that best describes a family name.</summary>
    internal static string ClassifyGenericFamily(string family) {
        string key = Normalize(family);
        if (IsMonospace(key)) return "monospace";
        if (IsSerif(key)) return "serif";
        return "sans-serif";
    }

    /// <summary>
    /// Pairs selected with compatibility intent. This policy does not qualify the installed font
    /// versions, vertical metrics, wrapping, or pagination; other pairs remain layout-sensitive.
    /// </summary>
    internal static bool IsIntendedCompatibleSubstitute(string source, string target) {
        string a = Normalize(source);
        string b = Normalize(target);
        return IsPair(a, b, "calibri", "carlito") || IsPair(a, b, "calibrilight", "carlito") ||
            IsPair(a, b, "cambria", "caladea") || IsPair(a, b, "georgia", "gelasio") ||
            ((a == "arial" || a == "helvetica") && (b == "liberationsans" || b == "arimo")) ||
            ((b == "arial" || b == "helvetica") && (a == "liberationsans" || a == "arimo")) ||
            ((a == "times" || a == "timesroman" || a == "timesnewroman") && (b == "liberationserif" || b == "tinos")) ||
            ((b == "times" || b == "timesroman" || b == "timesnewroman") && (a == "liberationserif" || a == "tinos")) ||
            ((a == "courier" || a == "couriernew") && (b == "liberationmono" || b == "cousine")) ||
            ((b == "courier" || b == "couriernew") && (a == "liberationmono" || a == "cousine"));
    }

    private static bool IsPair(string a, string b, string office, string open) =>
        (a == office && b == open) || (a == open && b == office);

    private static Dictionary<string, string[]> BuildChains() {
        var chains = new Dictionary<string, string[]>(StringComparer.Ordinal);
        // Theme families first: their fallback lists are the design source.
        foreach (OfficeTheme theme in new[] { OfficeTheme.Modern, OfficeTheme.Classic, OfficeTheme.Minimal, OfficeTheme.Boardroom, OfficeTheme.Dark }) {
            foreach (OfficeThemeFont font in new[] { theme.Fonts.Heading, theme.Fonts.Body, theme.ChartStyle.AxisFont, theme.ChartStyle.TitleFont }) {
                string key = Normalize(font.Family);
                if (chains.ContainsKey(key)) continue;
                var list = new List<string>();
                foreach (string fallback in font.Fallbacks) {
                    if (!IsGenericKeyword(Normalize(fallback))) list.Add(fallback);
                }
                chains[key] = list.ToArray();
            }
        }

        // Aliases of a theme family try that family first, then its theme fallback list.
        string[] aptos = Prepend("Aptos", chains["aptos"]);
        string[] calibri = Prepend("Calibri", chains.TryGetValue("calibri", out string[]? c) ? c : new[] { "Carlito", "Aptos", "Helvetica Neue", "Arial", "Liberation Sans" });
        string[] arial = Prepend("Arial", chains["arial"]);
        Add(chains, aptos, "aptosdisplay", "aptosbody", "aptosnarrow", "segoeui", "inter", "poppins", "montserrat", "opensans", "roboto", "lato", "sourcesanspro");
        Add(chains, calibri, "calibrilight");
        Add(chains, new[] { "Calibri", "Aptos", "Helvetica Neue", "Arial", "Liberation Sans" }, "carlito");
        Add(chains, new[] { "Cambria", "Georgia", "Liberation Serif", "Times New Roman" }, "caladea");
        Add(chains, new[] { "Georgia", "Cambria", "Caladea", "Liberation Serif", "Times New Roman" }, "gelasio");
        Add(chains, arial, "helvetica");
        Add(chains, new[] { "Arial", "Helvetica", "Liberation Sans", "Arimo", "Aptos", "Calibri" }, "helveticaneue");
        Add(chains, new[] { "Arial", "Arimo", "Helvetica Neue", "Helvetica" }, "liberationsans");
        Add(chains, new[] { "Arial", "Liberation Sans", "Helvetica Neue", "Helvetica" }, "arimo");
        Add(chains, new[] { "Times New Roman", "Liberation Serif", "Tinos", "Georgia", "DejaVu Serif" }, "times", "timesroman");
        Add(chains, new[] { "Liberation Serif", "Tinos", "Times", "Georgia", "DejaVu Serif" }, "timesnewroman");
        Add(chains, new[] { "Times New Roman", "Tinos", "Georgia", "DejaVu Serif" }, "liberationserif");
        Add(chains, new[] { "Times New Roman", "Liberation Serif", "Georgia", "DejaVu Serif" }, "tinos");
        Add(chains, new[] { "Liberation Mono", "Cousine", "Courier", "Consolas", "Menlo", "DejaVu Sans Mono" }, "couriernew");
        Add(chains, new[] { "Courier New", "Liberation Mono", "Cousine", "Consolas", "Menlo", "DejaVu Sans Mono" }, "courier", "liberationmono", "cousine");
        Add(chains, MonospaceChain, "consolas", "menlo", "monaco", "sfmono", "cascadiacode", "cascadiamono", "dejavusansmono");
        return chains;
    }

    private static string[] Prepend(string family, string[] chain) {
        var list = new List<string>(chain.Length + 1) { family };
        list.AddRange(chain);
        return list.ToArray();
    }

    private static void Add(Dictionary<string, string[]> chains, string[] chain, params string[] keys) {
        foreach (string key in keys) {
            if (!chains.ContainsKey(key)) chains[key] = chain;
        }
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

    // Symbol and dingbat families carry glyphs at private code points; a text face cannot substitute them.
    private static bool IsSymbolFamily(string key) =>
        key == "symbol" || key.StartsWith("wingdings", StringComparison.Ordinal) || key == "webdings" ||
        key == "marlett" || key == "mtextra" || key.Contains("dingbat") || key.Contains("emoji") ||
        key.Contains("math") || key == "segoeuisymbol" || key == "segoemdl2assets" || key == "segoefluenticons";

    private static bool IsGenericKeyword(string key) =>
        key == "sansserif" || key == "serif" || key == "monospace" || key == "cursive" || key == "fantasy" || key == "systemui";

    private static string Normalize(string family) =>
        family.Trim().Trim('"', '\'').Replace(" ", string.Empty).Replace("-", string.Empty).ToLowerInvariant();
}
