using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static OfficeTextFeatureSettings ResolveTextFeatureSettings(HtmlComputedStyle computed) {
        var features = new Dictionary<string, int>(StringComparer.Ordinal);

        string shorthand = computed.GetValue("font-variant");
        ApplyLigatures(shorthand, features);
        ApplyNumeric(shorthand, features);
        ApplyEastAsian(shorthand, features);
        ApplyKerning(computed.GetValue("font-kerning"), features);
        ApplyLigatures(computed.GetValue("font-variant-ligatures"), features);
        ApplyNumeric(computed.GetValue("font-variant-numeric"), features);
        ApplyEastAsian(computed.GetValue("font-variant-east-asian"), features);
        ApplyFeatureSettings(computed.GetValue("font-feature-settings"), features);
        return features.Count == 0 ? OfficeTextFeatureSettings.Default : new OfficeTextFeatureSettings(features);
    }

    private static string ResolveInheritedKeyword(string value, string? inherited, string fallback) {
        string normalized = value.Trim().ToLowerInvariant();
        return normalized.Length == 0 ? inherited ?? fallback : normalized;
    }

    private static void ApplyKerning(string value, IDictionary<string, int> features) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized == "none") features["kern"] = 0;
        else if (normalized == "normal") features["kern"] = 1;
    }

    private static void ApplyLigatures(string value, IDictionary<string, int> features) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "normal") return;
        if (normalized == "none") {
            Set(features, 0, "liga", "clig", "dlig", "hlig", "calt");
            return;
        }
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(normalized)) {
            switch (token) {
                case "common-ligatures": Set(features, 1, "liga", "clig"); break;
                case "no-common-ligatures": Set(features, 0, "liga", "clig"); break;
                case "discretionary-ligatures": features["dlig"] = 1; break;
                case "no-discretionary-ligatures": features["dlig"] = 0; break;
                case "historical-ligatures": features["hlig"] = 1; break;
                case "no-historical-ligatures": features["hlig"] = 0; break;
                case "contextual": features["calt"] = 1; break;
                case "no-contextual": features["calt"] = 0; break;
            }
        }
    }

    private static void ApplyNumeric(string value, IDictionary<string, int> features) {
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(value.Trim().ToLowerInvariant())) {
            switch (token) {
                case "lining-nums": Select(features, "lnum", "onum"); break;
                case "oldstyle-nums": Select(features, "onum", "lnum"); break;
                case "proportional-nums": Select(features, "pnum", "tnum"); break;
                case "tabular-nums": Select(features, "tnum", "pnum"); break;
                case "diagonal-fractions": Select(features, "frac", "afrc"); break;
                case "stacked-fractions": Select(features, "afrc", "frac"); break;
                case "ordinal": features["ordn"] = 1; break;
                case "slashed-zero": features["zero"] = 1; break;
            }
        }
    }

    private static void ApplyEastAsian(string value, IDictionary<string, int> features) {
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(value.Trim().ToLowerInvariant())) {
            switch (token) {
                case "jis78": Select(features, "jp78", "jp83", "jp90", "jp04"); break;
                case "jis83": Select(features, "jp83", "jp78", "jp90", "jp04"); break;
                case "jis90": Select(features, "jp90", "jp78", "jp83", "jp04"); break;
                case "jis04": Select(features, "jp04", "jp78", "jp83", "jp90"); break;
                case "simplified": Select(features, "smpl", "trad"); break;
                case "traditional": Select(features, "trad", "smpl"); break;
                case "full-width": Select(features, "fwid", "pwid"); break;
                case "proportional-width": Select(features, "pwid", "fwid"); break;
                case "ruby": features["ruby"] = 1; break;
            }
        }
    }

    private static void ApplyFeatureSettings(string value, IDictionary<string, int> features) {
        string normalized = value.Trim();
        if (normalized.Length == 0 || string.Equals(normalized, "normal", StringComparison.OrdinalIgnoreCase)) return;
        IReadOnlyList<string> entries = HtmlRenderCssValues.SplitTopLevelCommas(normalized);
        if (entries.Count > 128) return;
        foreach (string entry in entries) {
            IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(entry);
            if (tokens.Count < 1 || tokens.Count > 2 || !HtmlCounterStyleFormatter.TryUnquote(tokens[0], out string tag) || tag.Length != 4) continue;
            int setting = 1;
            if (tokens.Count == 2) {
                if (string.Equals(tokens[1], "on", StringComparison.OrdinalIgnoreCase)) setting = 1;
                else if (string.Equals(tokens[1], "off", StringComparison.OrdinalIgnoreCase)) setting = 0;
                else if (!int.TryParse(tokens[1], NumberStyles.None, CultureInfo.InvariantCulture, out setting) || setting < 0 || setting > ushort.MaxValue) continue;
            }
            bool printable = true;
            for (int index = 0; index < tag.Length; index++) printable &= tag[index] >= 0x20 && tag[index] <= 0x7E;
            if (printable) features[tag] = setting;
        }
    }

    private static void Set(IDictionary<string, int> features, int value, params string[] tags) {
        foreach (string tag in tags) features[tag] = value;
    }

    private static void Select(IDictionary<string, int> features, string enabled, params string[] disabled) {
        features[enabled] = 1;
        foreach (string tag in disabled) features[tag] = 0;
    }
}
