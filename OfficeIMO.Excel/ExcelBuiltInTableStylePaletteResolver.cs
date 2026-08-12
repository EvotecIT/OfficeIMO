namespace OfficeIMO.Excel {
    /// <summary>
    /// Resolves the visual palette of built-in Excel table styles for non-Excel renderers.
    /// </summary>
    internal static class ExcelBuiltInTableStylePaletteResolver {
        internal static bool TryCreate(
            ExcelDocument document,
            string? styleName,
            out ExcelBuiltInTableStylePalette? palette) {
            palette = null;
            if (!TryParse(styleName, out string? family, out int index)) {
                return false;
            }

            string light = ResolveThemeRgb(document, 0U, null, "FFFFFF");
            string dark = ResolveThemeRgb(document, 1U, null, "000000");
            int familyIndex = (index - 1) % 7;
            string baseColor = ResolveFamilyBaseColor(document, familyIndex);
            string paleColor = ResolveFamilyTintColor(document, familyIndex, 0.8D);
            string stripeColor = ResolveFamilyTintColor(document, familyIndex, 0.6D);
            string mutedColor = ResolveFamilyTintColor(document, familyIndex, -0.25D);

            switch (family) {
                case "Light":
                    if (index <= 7) {
                        palette = new ExcelBuiltInTableStylePalette(
                            headerFill: null,
                            headerText: familyIndex == 0 ? dark : mutedColor,
                            bodyFill: null,
                            stripeFill: paleColor,
                            bodyText: familyIndex == 0 ? dark : mutedColor,
                            border: familyIndex == 0 ? ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6") : baseColor,
                            headerBold: true);
                    } else if (index <= 14) {
                        palette = new ExcelBuiltInTableStylePalette(baseColor, light, null, null, dark, baseColor, headerBold: true);
                    } else {
                        palette = new ExcelBuiltInTableStylePalette(
                            headerFill: null,
                            headerText: dark,
                            bodyFill: null,
                            stripeFill: paleColor,
                            bodyText: dark,
                            border: familyIndex == 0 ? ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6") : baseColor,
                            headerBold: true);
                    }
                    return true;

                case "Medium":
                    if (index <= 7) {
                        palette = new ExcelBuiltInTableStylePalette(baseColor, light, null, paleColor, dark, baseColor, headerBold: true);
                    } else if (index <= 14) {
                        palette = new ExcelBuiltInTableStylePalette(baseColor, light, paleColor, stripeColor, dark, baseColor, headerBold: true);
                    } else if (index <= 21) {
                        string neutralStripe = ResolveThemeRgb(document, 0U, -0.15D, "D9D9D9");
                        palette = new ExcelBuiltInTableStylePalette(baseColor, light, null, neutralStripe, dark, baseColor, headerBold: true);
                    } else {
                        palette = new ExcelBuiltInTableStylePalette(paleColor, dark, paleColor, stripeColor, dark, baseColor, headerBold: true);
                    }
                    return true;

                case "Dark":
                    if (index <= 7) {
                        string bodyFill = familyIndex == 0
                            ? ResolveThemeRgb(document, 1U, 0.45D, "737373")
                            : baseColor;
                        string bodyStripe = familyIndex == 0
                            ? ResolveThemeRgb(document, 1U, 0.25D, "404040")
                            : mutedColor;
                        palette = new ExcelBuiltInTableStylePalette(dark, light, bodyFill, bodyStripe, light, baseColor, headerBold: true);
                    } else if (index == 8) {
                        palette = new ExcelBuiltInTableStylePalette(
                            dark,
                            light,
                            ResolveThemeRgb(document, 0U, -0.15D, "D9D9D9"),
                            ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6"),
                            dark,
                            ResolveThemeRgb(document, 0U, -0.35D, "A6A6A6"),
                            headerBold: false);
                    } else {
                        uint headerTheme = index == 9 ? 5U : index == 10 ? 7U : 9U;
                        uint bodyTheme = index == 9 ? 4U : index == 10 ? 6U : 8U;
                        string header = ResolveThemeRgb(document, headerTheme, null, index == 9 ? "E97132" : index == 10 ? "0F9ED5" : "4EA72E");
                        string body = ResolveThemeRgb(document, bodyTheme, 0.8D, index == 9 ? "C0E6F5" : index == 10 ? "C1F0C8" : "F2CEEF");
                        string stripe = ResolveThemeRgb(document, bodyTheme, 0.6D, index == 9 ? "83CCEB" : index == 10 ? "83E28E" : "E49EDD");
                        palette = new ExcelBuiltInTableStylePalette(header, light, body, stripe, dark, header, headerBold: false);
                    }
                    return true;
            }

            return false;
        }

        private static bool TryParse(string? styleName, out string? family, out int index) {
            family = null;
            index = 0;
            if (string.IsNullOrWhiteSpace(styleName) ||
                !styleName!.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string suffix = styleName.Substring("TableStyle".Length);
            foreach (string candidate in new[] { "Light", "Medium", "Dark" }) {
                if (!suffix.StartsWith(candidate, StringComparison.OrdinalIgnoreCase) ||
                    !int.TryParse(suffix.Substring(candidate.Length), out int parsed)) {
                    continue;
                }

                int maximum = candidate == "Light" ? 21 : candidate == "Medium" ? 28 : 11;
                if (parsed < 1 || parsed > maximum) {
                    return false;
                }

                family = candidate;
                index = parsed;
                return true;
            }

            return false;
        }

        private static string ResolveFamilyBaseColor(ExcelDocument document, int familyIndex) {
            if (familyIndex == 0) {
                return ResolveThemeRgb(document, 1U, null, "000000");
            }

            string[] fallbacks = { "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E" };
            return ResolveThemeRgb(document, (uint)(familyIndex + 3), null, fallbacks[familyIndex - 1]);
        }

        private static string ResolveFamilyTintColor(ExcelDocument document, int familyIndex, double tint) {
            if (familyIndex == 0) {
                string neutralFallback = tint >= 0D
                    ? tint >= 0.7D ? "D9D9D9" : tint >= 0.5D ? "A6A6A6" : "737373"
                    : "A6A6A6";
                double lightTint = tint >= 0D ? tint - 0.95D : tint;
                return ResolveThemeRgb(document, 0U, lightTint, neutralFallback);
            }

            string[] paleFallbacks = { "C0E6F5", "FBE2D5", "C1F0C8", "CAEDFB", "F2CEEF", "DAF2D0" };
            string[] stripeFallbacks = { "83CCEB", "F7C7AC", "83E28E", "94DCF8", "E49EDD", "B5E6A2" };
            string[] mutedFallbacks = { "104861", "BE5014", "12501A", "0C769E", "782170", "3C7D22" };
            string fallback = tint >= 0.7D
                ? paleFallbacks[familyIndex - 1]
                : tint >= 0D
                    ? stripeFallbacks[familyIndex - 1]
                    : mutedFallbacks[familyIndex - 1];
            return ResolveThemeRgb(document, (uint)(familyIndex + 3), tint, fallback);
        }

        private static string ResolveThemeRgb(ExcelDocument document, uint themeIndex, double? tint, string fallback) {
            string? argb = document.ResolveThemeColorArgb(themeIndex, tint);
            if (string.IsNullOrWhiteSpace(argb)) {
                return fallback;
            }

            string normalized = argb!.Trim().TrimStart('#');
            return normalized.Length == 8 ? normalized.Substring(2) : normalized.Length == 6 ? normalized : fallback;
        }
    }

    internal sealed class ExcelBuiltInTableStylePalette {
        internal ExcelBuiltInTableStylePalette(
            string? headerFill,
            string? headerText,
            string? bodyFill,
            string? stripeFill,
            string? bodyText,
            string? border,
            bool headerBold) {
            HeaderFill = headerFill;
            HeaderText = headerText;
            BodyFill = bodyFill;
            StripeFill = stripeFill;
            BodyText = bodyText;
            Border = border;
            HeaderBold = headerBold;
        }

        internal string? HeaderFill { get; }
        internal string? HeaderText { get; }
        internal string? BodyFill { get; }
        internal string? StripeFill { get; }
        internal string? BodyText { get; }
        internal string? Border { get; }
        internal bool HeaderBold { get; }
    }
}
