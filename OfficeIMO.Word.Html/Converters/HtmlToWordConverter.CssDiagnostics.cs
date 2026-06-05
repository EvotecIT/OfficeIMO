using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly HashSet<string> _supportedCssDiagnosticProperties = new(StringComparer.OrdinalIgnoreCase) {
            "background-color",
            "border",
            "border-collapse",
            "border-spacing",
            "break-after",
            "break-before",
            "color",
            "direction",
            "float",
            "font",
            "font-family",
            "font-size",
            "font-style",
            "font-variant",
            "font-weight",
            "height",
            "letter-spacing",
            "line-height",
            "list-style",
            "list-style-type",
            "margin",
            "margin-bottom",
            "margin-left",
            "margin-right",
            "margin-top",
            "padding",
            "padding-bottom",
            "padding-left",
            "padding-right",
            "padding-top",
            "page-break-after",
            "page-break-before",
            "text-align",
            "text-decoration",
            "text-decoration-line",
            "text-decoration-style",
            "text-indent",
            "text-transform",
            "vertical-align",
            "white-space",
            "width",
        };

        private void ReportUnsupportedInlineCssDiagnostics(IElement element) {
            var style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) {
                return;
            }

            var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rawPropertyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in ParseCssDeclarationPairs(style!)) {
                declarations[pair.Key] = pair.Value;
                rawPropertyNames.Add(pair.Key);
            }

            var hasRawFontShorthand = rawPropertyNames.Contains("font");
            foreach (var property in ParseInlineDeclaration(style)) {
                if (rawPropertyNames.Contains(property.Name)) {
                    continue;
                }
                if (hasRawFontShorthand && property.Name.StartsWith("font-", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                declarations[property.Name] = property.Value;
            }

            ReportUnsupportedCssDiagnostics(element, declarations);
        }

        private void ReportUnsupportedCssDiagnostics(IElement element, IEnumerable<string> propertyNames) {
            var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var propertyName in propertyNames) {
                if (!string.IsNullOrWhiteSpace(propertyName)) {
                    declarations[propertyName] = string.Empty;
                }
            }

            ReportUnsupportedCssDiagnostics(element, declarations);
        }

        private void ReportUnsupportedCssDiagnostics(IElement element, IReadOnlyDictionary<string, string> declarations) {
            foreach (var pair in declarations) {
                var propertyName = pair.Key;
                if (string.IsNullOrWhiteSpace(propertyName)) {
                    continue;
                }

                var elementName = element.TagName.ToLowerInvariant();
                if (!IsSupportedCssDiagnosticProperty(propertyName)) {
                    AddUnsupportedCssDiagnostic(
                        "UnsupportedCssDeclaration",
                        "CSS declaration is not currently mapped to Word output.",
                        $"{elementName}:{propertyName}");
                    continue;
                }

                if (TryGetUnsupportedCssValueReason(elementName, propertyName, pair.Value, out var reason)) {
                    AddUnsupportedCssDiagnostic(
                        "UnsupportedCssValue",
                        "CSS declaration value is not currently mapped to Word output.",
                        $"{elementName}:{propertyName}",
                        reason);
                }
            }
        }

        private static bool IsSupportedCssDiagnosticProperty(string propertyName) =>
            _supportedCssDiagnosticProperties.Contains(propertyName) ||
            IsSupportedBorderSideShorthand(propertyName);

        private void AddUnsupportedCssDiagnostic(string code, string message, string source, string? detail = null) {
            if (_options.UnsupportedCssHandling == HtmlUnsupportedCssHandling.Ignore) {
                return;
            }

            var key = $"{code}|{source}|{detail}";
            if (!_unsupportedCssDiagnosticKeys.Add(key)) {
                return;
            }

            if (_options.UnsupportedCssHandling == HtmlUnsupportedCssHandling.Error) {
                var exception = new HtmlUnsupportedCssException(code, message, source, detail);
                AddDiagnostic(_options, code, message, source, exception, HtmlConversionDiagnosticSeverity.Error);
                throw exception;
            }

            AddDiagnostic(_options, code, message, source, detail == null ? null : new HtmlUnsupportedCssException(code, message, source, detail));
        }

        private static bool TryGetUnsupportedCssValueReason(string elementName, string propertyName, string? rawValue, out string reason) {
            reason = string.Empty;
            var value = NormalizeCssDiagnosticValue(rawValue);
            if (string.IsNullOrWhiteSpace(value) || IsCssGlobalKeyword(value)) {
                return false;
            }

            var lower = value.ToLowerInvariant();
            switch (propertyName.ToLowerInvariant()) {
                case "color":
                    return TryUnsupportedColorValue(value, "color", out reason);
                case "background-color":
                    return TryUnsupportedColorValue(value, "background color", out reason);
                case "font":
                    if (!IsSupportedFontShorthand(value)) {
                        reason = $"Unsupported font value '{value}'.";
                        return true;
                    }
                    return false;
                case "font-size":
                    if (!TryParseFontSize(value, out _)) {
                        reason = $"Unsupported font-size value '{value}'.";
                        return true;
                    }
                    return false;
                case "font-weight":
                    if (lower == "normal" || lower == "bold" || int.TryParse(lower, out _)) {
                        return false;
                    }
                    reason = $"Unsupported font-weight value '{value}'.";
                    return true;
                case "font-style":
                    if (lower is "normal" or "italic" or "oblique") {
                        return false;
                    }
                    reason = $"Unsupported font-style value '{value}'.";
                    return true;
                case "font-variant":
                    if (lower is "normal" or "small-caps") {
                        return false;
                    }
                    reason = $"Unsupported font-variant value '{value}'.";
                    return true;
                case "line-height":
                    if (!TryParseSupportedLineHeight(value)) {
                        reason = $"Unsupported line-height value '{value}'.";
                        return true;
                    }
                    return false;
                case "text-align":
                    if (lower is "left" or "right" or "center" or "justify" or "start" or "end") {
                        return false;
                    }
                    reason = $"Unsupported text-align value '{value}'.";
                    return true;
                case "text-decoration":
                    if (TryGetUnsupportedTextDecorationReason(value, "text-decoration", out reason)) {
                        return true;
                    }
                    return false;
                case "text-decoration-line":
                    if (TryGetUnsupportedTextDecorationLineReason(value, "text-decoration-line", out reason)) {
                        return true;
                    }
                    return false;
                case "text-decoration-style":
                    if (!IsSupportedTextDecorationStyle(lower)) {
                        reason = $"Unsupported text-decoration-style value '{value}'.";
                        return true;
                    }
                    return false;
                case "direction":
                    if (lower is "ltr" or "rtl") {
                        return false;
                    }
                    reason = $"Unsupported direction value '{value}'.";
                    return true;
                case "text-transform":
                    if (lower is "none" or "uppercase" or "lowercase" or "capitalize") {
                        return false;
                    }
                    reason = $"Unsupported text-transform value '{value}'.";
                    return true;
                case "vertical-align":
                    if (IsTableCellElement(elementName) && lower is "top" or "middle" or "center" or "bottom") {
                        return false;
                    }
                    if (lower is "baseline" or "super" or "sup" or "sub") {
                        return false;
                    }
                    reason = $"Unsupported vertical-align value '{value}'.";
                    return true;
                case "white-space":
                    if (lower is "normal" or "pre" or "pre-wrap" or "nowrap") {
                        return false;
                    }
                    reason = $"Unsupported white-space value '{value}'.";
                    return true;
                case "list-style":
                    if (ExtractListStyleToken(value) == null) {
                        reason = $"Unsupported list-style value '{value}'.";
                        return true;
                    }
                    return false;
                case "list-style-type":
                    if (NormalizeListStyleToken(value) == null) {
                        reason = $"Unsupported list-style-type value '{value}'.";
                        return true;
                    }
                    return false;
                case "break-before":
                case "break-after":
                    if (lower is "auto" or "avoid" or "page" or "always" or "left" or "right") {
                        return false;
                    }
                    reason = $"Unsupported {propertyName} value '{value}'.";
                    return true;
                case "page-break-before":
                case "page-break-after":
                    if (lower is "auto" or "avoid" or "always" or "left" or "right") {
                        return false;
                    }
                    reason = $"Unsupported {propertyName} value '{value}'.";
                    return true;
                case "float":
                    if (lower is "none" or "left" or "right") {
                        return false;
                    }
                    reason = $"Unsupported float value '{value}'.";
                    return true;
                case "border":
                case "border-left":
                case "border-right":
                case "border-top":
                case "border-bottom":
                    if (!TryParseBorder(value, out _, out _, out _)) {
                        reason = $"Unsupported {propertyName} value '{value}'.";
                        return true;
                    }
                    return false;
                case "border-collapse":
                    if (lower is "collapse" or "separate") {
                        return false;
                    }
                    reason = $"Unsupported border-collapse value '{value}'.";
                    return true;
                case "border-spacing":
                    if (!IsSupportedBorderSpacingValue(value)) {
                        reason = $"Unsupported border-spacing value '{value}'.";
                        return true;
                    }
                    return false;
                case "letter-spacing":
                case "text-indent":
                    if (!IsSupportedCssLength(value, allowNegative: true, allowPercent: false, allowAuto: false)) {
                        reason = $"Unsupported {propertyName} value '{value}'.";
                        return true;
                    }
                    return false;
            }

            if (propertyName.StartsWith("margin", StringComparison.OrdinalIgnoreCase) ||
                propertyName.StartsWith("padding", StringComparison.OrdinalIgnoreCase)) {
                if (!IsSupportedBoxLengthList(value, allowAuto: propertyName.StartsWith("margin", StringComparison.OrdinalIgnoreCase))) {
                    reason = $"Unsupported {propertyName} value '{value}'.";
                    return true;
                }
                return false;
            }

            if (propertyName.Equals("width", StringComparison.OrdinalIgnoreCase) ||
                propertyName.Equals("height", StringComparison.OrdinalIgnoreCase)) {
                if (!IsSupportedCssLength(value, allowNegative: false, allowPercent: propertyName.Equals("width", StringComparison.OrdinalIgnoreCase), allowAuto: true)) {
                    reason = $"Unsupported {propertyName} value '{value}'.";
                    return true;
                }
            }

            return false;
        }

        private static bool IsTableCellElement(string elementName) =>
            elementName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
            elementName.Equals("th", StringComparison.OrdinalIgnoreCase);

        private static bool IsSupportedBorderSpacingValue(string value) {
            var tokens = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length is < 1 or > 2) {
                return false;
            }

            foreach (var token in tokens) {
                if (!IsSupportedCssLength(token, allowNegative: false, allowPercent: false, allowAuto: false)) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsSupportedFontShorthand(string value) {
            var tokens = TokenizeFontShorthand(value);
            if (tokens.Count == 0) {
                return false;
            }

            int sizeIndex = -1;
            for (int i = 0; i < tokens.Count; i++) {
                var token = tokens[i];
                if (token.IndexOf('/') >= 0) {
                    return false;
                }

                if (TryParseFontSize(token, out _)) {
                    sizeIndex = i;
                    break;
                }
            }

            if (sizeIndex < 0 || sizeIndex + 1 >= tokens.Count) {
                return false;
            }

            for (int i = 0; i < sizeIndex; i++) {
                var token = tokens[i].Trim().ToLowerInvariant();
                if (token is "normal" or "italic" or "oblique" or "small-caps" or "bold" or "bolder" or "lighter") {
                    continue;
                }
                if (int.TryParse(token, out _)) {
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool IsSupportedBorderSideShorthand(string propertyName) =>
            propertyName.Equals("border-left", StringComparison.OrdinalIgnoreCase) ||
            propertyName.Equals("border-right", StringComparison.OrdinalIgnoreCase) ||
            propertyName.Equals("border-top", StringComparison.OrdinalIgnoreCase) ||
            propertyName.Equals("border-bottom", StringComparison.OrdinalIgnoreCase);

        private static bool TryUnsupportedColorValue(string value, string label, out string reason) {
            reason = string.Empty;
            if (value.Equals("transparent", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("currentColor", StringComparison.OrdinalIgnoreCase)) {
                reason = $"Unsupported {label} value '{value}'.";
                return true;
            }

            if (NormalizeColor(value) != null) {
                return false;
            }

            reason = $"Unsupported {label} value '{value}'.";
            return true;
        }

        private static bool TryParseSupportedLineHeight(string value) {
            var lower = value.Trim().ToLowerInvariant();
            if (lower == "normal") {
                return true;
            }

            var parsed = CssStyleMapper.ParseStyles($"line-height:{value}");
            return parsed.LineHeight.HasValue;
        }

        private static bool TryGetUnsupportedTextDecorationReason(string value, string propertyName, out string reason) {
            reason = string.Empty;
            var tokens = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0) {
                return false;
            }

            foreach (var token in tokens) {
                var lower = token.Trim().ToLowerInvariant();
                if (IsSupportedTextDecorationLine(lower) || IsSupportedTextDecorationStyle(lower)) {
                    continue;
                }

                if (NormalizeColor(token) != null ||
                    lower is "currentcolor" or "transparent") {
                    reason = $"Unsupported {propertyName} color token '{token}' in value '{value}'.";
                    return true;
                }

                reason = $"Unsupported {propertyName} token '{token}' in value '{value}'.";
                return true;
            }

            return false;
        }

        private static bool TryGetUnsupportedTextDecorationLineReason(string value, string propertyName, out string reason) {
            reason = string.Empty;
            foreach (var token in value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                var lower = token.Trim().ToLowerInvariant();
                if (IsSupportedTextDecorationLine(lower)) {
                    continue;
                }

                reason = $"Unsupported {propertyName} token '{token}' in value '{value}'.";
                return true;
            }

            return false;
        }

        private static bool IsSupportedTextDecorationLine(string value) =>
            value is "none" or "underline" or "line-through";

        private static bool IsSupportedTextDecorationStyle(string value) =>
            value is "solid" or "double" or "dotted" or "dashed" or "wavy";

        private static bool IsSupportedBoxLengthList(string value, bool allowAuto) {
            var tokens = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0 || tokens.Length > 4) {
                return false;
            }

            foreach (var token in tokens) {
                if (!IsSupportedCssLength(token, allowNegative: true, allowPercent: false, allowAuto: allowAuto)) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsSupportedCssLength(string value, bool allowNegative, bool allowPercent, bool allowAuto) {
            var lower = value.Trim().ToLowerInvariant();
            if (allowAuto && lower == "auto") {
                return true;
            }

            if (IsCssZeroLength(lower)) {
                return true;
            }

            if (allowPercent && lower.EndsWith("%", StringComparison.Ordinal) &&
                double.TryParse(lower.Substring(0, lower.Length - 1), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var percent)) {
                return percent >= 0;
            }

            var declaration = ParseInlineDeclaration($"x:{value}");
            var raw = declaration.GetProperty("x")?.RawValue;
            if (allowNegative ? TryConvertToTwipAllowNegative(raw, out _) : TryConvertToTwip(raw, out _)) {
                return true;
            }

            return IsSupportedCssLengthLiteral(lower, allowNegative);
        }

        private static bool IsSupportedCssLengthLiteral(string value, bool allowNegative) {
            string[] units = { "px", "pt", "em", "rem", "cm", "mm", "in", "pc", "q" };
            foreach (var unit in units) {
                if (!value.EndsWith(unit, StringComparison.Ordinal)) {
                    continue;
                }

                if (!double.TryParse(value.Substring(0, value.Length - unit.Length), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number)) {
                    return false;
                }

                return allowNegative || number >= 0;
            }

            return false;
        }

        private static bool IsCssZeroLength(string value) {
            if (value == "0") {
                return true;
            }

            string[] units = { "px", "pt", "em", "rem", "cm", "mm", "in", "pc", "q" };
            foreach (var unit in units) {
                if (value.EndsWith(unit, StringComparison.Ordinal) &&
                    double.TryParse(value.Substring(0, value.Length - unit.Length), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number)) {
                    return number == 0;
                }
            }

            return false;
        }

        private static bool IsCssGlobalKeyword(string value) {
            var lower = value.Trim().ToLowerInvariant();
            return lower is "inherit" or "initial" or "unset" or "revert" or "revert-layer";
        }

        private static string NormalizeCssDiagnosticValue(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            var normalized = value!.Trim();
            var importantIndex = normalized.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
            if (importantIndex >= 0) {
                normalized = normalized.Substring(0, importantIndex).Trim();
            }

            return normalized;
        }

        private static IEnumerable<KeyValuePair<string, string>> ParseCssDeclarationPairs(string style) {
            foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length == 2) {
                    var name = pieces[0].Trim();
                    if (!string.IsNullOrWhiteSpace(name)) {
                        yield return new KeyValuePair<string, string>(name, pieces[1].Trim());
                    }
                }
            }
        }
    }
}
