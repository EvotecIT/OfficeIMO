using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Values;
using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static bool TryResolveRootFontSizePixels(
            IReadOnlyDictionary<string, (string Value, Priority Specificity, bool Important, int Order)> declarations,
            out double pixels) {
            return TryResolveDeclaredFontSizePixels(
                GetWinningFontSizeDeclaration(declarations),
                _renderDevice.FontSize,
                _renderDevice.FontSize,
                out pixels);
        }

        private double ResolveComputedFontSizePixels(IElement? element) {
            if (element == null) {
                return _rootFontSizePixels;
            }
            if (_computedFontSizePixels.TryGetValue(element, out double cached)) {
                return cached;
            }

            CacheComputedFontSizePixels(
                element,
                CollectCssDeclarations(element, inheritedOnly: false));
            return _computedFontSizePixels[element];
        }

        private void CacheComputedFontSizePixels(
            IElement element,
            IReadOnlyDictionary<string, (string Value, Priority Specificity, bool Important, int Order)> declarations) {
            double inheritedFontSizePixels = ResolveComputedFontSizePixels(element.ParentElement);
            if (!TryResolveDeclaredFontSizePixels(
                    GetWinningFontSizeDeclaration(declarations),
                    inheritedFontSizePixels,
                    _rootFontSizePixels,
                    out double pixels)) {
                pixels = inheritedFontSizePixels;
            }

            _computedFontSizePixels[element] = pixels;
        }

        private static bool TryResolveDeclaredFontSizePixels(
            string? declaration,
            double inheritedFontSizePixels,
            double rootFontSizePixels,
            out double pixels) {
            if (TryResolveFontSizePixels(
                    declaration,
                    inheritedFontSizePixels,
                    rootFontSizePixels,
                    out pixels)) {
                return true;
            }

            foreach (string token in TokenizeFontShorthand(declaration ?? string.Empty)) {
                string sizeToken = token;
                int slashIndex = token.IndexOf('/');
                if (slashIndex >= 0) {
                    sizeToken = token.Substring(0, slashIndex);
                }
                if (TryResolveFontSizePixels(
                        sizeToken,
                        inheritedFontSizePixels,
                        rootFontSizePixels,
                        out pixels)) {
                    return true;
                }
            }

            pixels = 0d;
            return false;
        }

        private static string? GetWinningFontSizeDeclaration(
            IReadOnlyDictionary<string, (string Value, Priority Specificity, bool Important, int Order)> declarations) {
            bool hasFontSize = declarations.TryGetValue("font-size", out var fontSize);
            bool hasFont = declarations.TryGetValue("font", out var font);
            if (!hasFontSize) {
                return hasFont ? font.Value : null;
            }
            if (!hasFont) {
                return fontSize.Value;
            }

            if (font.Important != fontSize.Important) {
                return font.Important ? font.Value : fontSize.Value;
            }
            if (font.Specificity != fontSize.Specificity) {
                return font.Specificity > fontSize.Specificity ? font.Value : fontSize.Value;
            }
            return font.Order >= fontSize.Order ? font.Value : fontSize.Value;
        }

        private static bool TryResolveFontSizePixels(
            string? text,
            double inheritedFontSizePixels,
            double rootFontSizePixels,
            out double pixels) {
            pixels = 0d;
            if (string.IsNullOrWhiteSpace(text)) {
                return false;
            }

            string normalized = text!.Trim().ToLowerInvariant();
            if (normalized is "inherit" or "unset") {
                pixels = inheritedFontSizePixels;
                return true;
            }
            if (normalized is "initial" or "revert" or "revert-layer") {
                pixels = _renderDevice.FontSize;
                return true;
            }
            if (TryParseRelativeFontSize(normalized, "rem", rootFontSizePixels, out pixels) ||
                TryParseRelativeFontSize(normalized, "em", inheritedFontSizePixels, out pixels)) {
                return pixels > 0d;
            }
            if (normalized.EndsWith("%", StringComparison.Ordinal) &&
                double.TryParse(
                    normalized.Substring(0, normalized.Length - 1),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double percent)) {
                pixels = inheritedFontSizePixels * percent / 100d;
                return pixels > 0d;
            }

            var declaration = ParseInlineDeclaration($"font-size:{normalized}");
            if (declaration.GetProperty("font-size")?.RawValue is CssLengthValue length) {
                try {
                    pixels = length.ToPixel(_renderDevice);
                    return pixels > 0d && !double.IsInfinity(pixels) && !double.IsNaN(pixels);
                } catch {
                    return false;
                }
            }

            if (_namedFontSizes.TryGetValue(normalized, out int named)) {
                pixels = named;
                return true;
            }

            return false;
        }

        private static bool TryParseRelativeFontSize(
            string value,
            string unit,
            double basisPixels,
            out double pixels) {
            pixels = 0d;
            if (!value.EndsWith(unit, StringComparison.Ordinal) ||
                !double.TryParse(
                    value.Substring(0, value.Length - unit.Length),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double multiple)) {
                return false;
            }

            pixels = basisPixels * multiple;
            return true;
        }
    }
}
