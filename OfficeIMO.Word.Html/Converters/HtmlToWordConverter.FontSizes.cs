using AngleSharp.Css.Values;
using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static bool TryResolveRootFontSizePixels(string? text, out double pixels) {
            return TryResolveFontSizePixels(
                text,
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

            string? value = null;
            if (_cssRules.Count == 0) {
                if (TryGetInlineProperty(element.GetAttribute("style"), "font-size", out string inlineValue)) {
                    value = inlineValue;
                }
            } else {
                var declarations = CollectCssDeclarations(element, inheritedOnly: false);
                if (declarations.TryGetValue("font-size", out var declaration)) {
                    value = declaration.Value;
                }
            }

            CacheComputedFontSizePixels(element, value);
            return _computedFontSizePixels[element];
        }

        private void CacheComputedFontSizePixels(IElement element, string? value) {
            double inheritedFontSizePixels = ResolveComputedFontSizePixels(element.ParentElement);
            if (string.IsNullOrWhiteSpace(value) ||
                !TryResolveFontSizePixels(
                    value,
                    inheritedFontSizePixels,
                    _rootFontSizePixels,
                    out double pixels)) {
                pixels = inheritedFontSizePixels;
            }

            _computedFontSizePixels[element] = pixels;
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
