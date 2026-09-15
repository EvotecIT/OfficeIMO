using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumSvgCssRules = 4096;
    private const int MaximumSvgCssDeclarations = 32768;
    private const long MaximumSvgCssMatchWork = 2_000_000L;
    private const long MaximumSvgComputedCssCharacters = 16L * 1024L * 1024L;

    private static bool ApplySvgStylesheets(XElement root, ref int unsupported, bool requireCompleteCss = false) {
        var rules = new List<SvgCssRule>();
        int declarations = 0;
        bool complete = true;
        XNamespace svgNamespace = root.Name.Namespace;
        XElement selectorRoot = new XElement(root);
        foreach (XElement style in selectorRoot.DescendantsAndSelf().Where(element =>
                     IsNativeSvgElement(element, svgNamespace) &&
                     element.Name.LocalName.Equals("style", StringComparison.Ordinal))) {
            if (!IsSupportedSvgStylesheetElement(style)) {
                unsupported++;
                complete = false;
                continue;
            }
            if (rules.Count >= MaximumSvgCssRules || declarations >= MaximumSvgCssDeclarations) {
                unsupported++;
                complete = false;
                break;
            }
            if (!ParseSvgCssRules(RemoveSvgCssComments(style.Value), rules, ref declarations, ref unsupported)) {
                complete = false;
            }
        }
        var validationPaintServers = new SvgPaintServerRegistry(SvgDefinitionRegistry.Create(root));
        var validatedRules = new List<SvgCssRule>(rules.Count);
        foreach (SvgCssRule rule in rules) {
            var validDeclarations = new List<SvgCssDeclaration>(rule.Declarations.Count);
            foreach (SvgCssDeclaration declaration in rule.Declarations) {
                if (IsSupportedSvgCssDeclaration(declaration, validationPaintServers, ref unsupported)) {
                    validDeclarations.Add(declaration);
                } else if (IsUnsupportedSvgCssWideKeyword(declaration.Value) || requireCompleteCss) {
                    complete = false;
                }
            }
            validatedRules.Add(new SvgCssRule(rule.Selector, rule.Parts, validDeclarations.AsReadOnly(), rule.Specificity, rule.Order));
        }
        rules = validatedRules;

        XElement[] elements = root.DescendantsAndSelf().ToArray();
        XElement[] selectorElements = selectorRoot.DescendantsAndSelf().ToArray();
        long remainingMatchWork = MaximumSvgCssMatchWork;
        long remainingComputedCssCharacters = MaximumSvgComputedCssCharacters;
        bool matchWorkExceeded = false;
        var computedCustomProperties = new Dictionary<XElement, IReadOnlyDictionary<string, string>>();
        for (int elementIndex = 0; elementIndex < elements.Length; elementIndex++) {
            XElement element = elements[elementIndex];
            XElement selectorElement = selectorElements[elementIndex];
            if (!IsNativeSvgElement(element, svgNamespace) ||
                element.Name.LocalName.Equals("style", StringComparison.Ordinal)) continue;
            var winners = new Dictionary<string, SvgCssWinner>(StringComparer.Ordinal);
            foreach (SvgCssRule rule in rules) {
                if (!MatchesSvgSelector(
                        selectorElement,
                        rule.Parts,
                        svgNamespace,
                        ref remainingMatchWork,
                        ref matchWorkExceeded)) {
                    if (matchWorkExceeded) {
                        unsupported++;
                        return false;
                    }
                    continue;
                }
                foreach (SvgCssDeclaration declaration in rule.Declarations) {
                    SetSvgCssWinner(winners, declaration, rule.Specificity, rule.Order);
                }
            }
            int inlineOrder = rules.Count + 1;
            bool hadInlineStyle = !string.IsNullOrWhiteSpace(selectorElement.Attribute("style")?.Value);
            int inlineDeclarations = 0;
            IReadOnlyList<SvgCssDeclaration> parsedInline = ParseSvgCssDeclarations(
                selectorElement.Attribute("style")?.Value,
                ref inlineDeclarations,
                out bool inlineComplete);
            if (!inlineComplete) {
                unsupported++;
                complete = false;
            }
            foreach (SvgCssDeclaration declaration in parsedInline) {
                if (!IsSupportedSvgCssDeclaration(declaration, validationPaintServers, ref unsupported)) {
                    if (IsUnsupportedSvgCssWideKeyword(declaration.Value) || requireCompleteCss) complete = false;
                    continue;
                }
                SetSvgCssWinner(winners, declaration, SvgCssSpecificity.Inline, inlineOrder++);
            }
            if (!TryResolveSvgCustomProperties(
                element,
                winners,
                computedCustomProperties,
                ref remainingComputedCssCharacters,
                out IReadOnlyDictionary<string, string> customProperties)) {
                unsupported++;
                return false;
            }
            computedCustomProperties[element] = customProperties;
            if (winners.Count == 0) {
                if (hadInlineStyle) element.SetAttributeValue("style", null);
                continue;
            }

            var css = new StringBuilder();
            var appliedPresentationProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, SvgCssWinner> pair in winners.OrderBy(item => item.Value.Order)) {
                if (pair.Key.StartsWith("--", StringComparison.Ordinal)) continue;
                if (IsSvgPresentationPropertyName(pair.Key)) appliedPresentationProperties.Add(pair.Key);
                if (!TryResolveSvgComputedCssValue(
                        pair.Key,
                        pair.Value.Value,
                        customProperties,
                        element.Parent,
                        validationPaintServers,
                        ref unsupported,
                        out string? value)) {
                    unsupported++;
                    complete = false;
                    continue;
                }
                if (value == null) continue;
                long declarationCharacters = pair.Key.Length + value.Length + (css.Length > 0 ? 2L : 1L);
                if (declarationCharacters > remainingComputedCssCharacters) {
                    unsupported++;
                    return false;
                }
                remainingComputedCssCharacters -= declarationCharacters;
                if (css.Length > 0) css.Append(';');
                css.Append(pair.Key).Append(':').Append(value);
            }
            element.SetAttributeValue("style", css.Length == 0 ? null : css.ToString());
            foreach (string propertyName in appliedPresentationProperties) {
                element.Attribute(propertyName)?.Remove();
            }
        }
        return complete;
    }

    private static bool TryResolveSvgCustomProperties(
        XElement element,
        IReadOnlyDictionary<string, SvgCssWinner> winners,
        IReadOnlyDictionary<XElement, IReadOnlyDictionary<string, string>> computedProperties,
        ref long remainingComputedCssCharacters,
        out IReadOnlyDictionary<string, string> properties) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        XElement? parent = element.Parent;
        if (parent != null && computedProperties.TryGetValue(parent, out IReadOnlyDictionary<string, string>? inherited)) {
            foreach (KeyValuePair<string, string> property in inherited) {
                if (!TryConsumeSvgComputedCssCharacters(property.Key, property.Value, ref remainingComputedCssCharacters)) {
                    properties = result;
                    return false;
                }
                result[property.Key] = property.Value;
            }
        }
        foreach (KeyValuePair<string, SvgCssWinner> winner in winners) {
            if (!winner.Key.StartsWith("--", StringComparison.Ordinal)) continue;
            string normalized = winner.Value.Value.Trim();
            if (normalized.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("unset", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }
            if (normalized.Equals("initial", StringComparison.OrdinalIgnoreCase)) {
                result.Remove(winner.Key);
                continue;
            }
            if (!TryConsumeSvgComputedCssCharacters(winner.Key, winner.Value.Value, ref remainingComputedCssCharacters)) {
                properties = result;
                return false;
            }
            result[winner.Key] = winner.Value.Value;
        }
        properties = result;
        return true;
    }

    private static bool TryConsumeSvgComputedCssCharacters(
        string name,
        string value,
        ref long remainingCharacters) {
        long required = name.Length + value.Length + 2L;
        if (required > remainingCharacters) return false;
        remainingCharacters -= required;
        return true;
    }

    private static bool IsSupportedSvgStylesheetElement(XElement style) {
        string? media = style.Attribute("media")?.Value;
        if (!string.IsNullOrWhiteSpace(media) && !media!.Trim().Equals("all", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        string? type = style.Attribute("type")?.Value;
        if (!string.IsNullOrWhiteSpace(type) && !type!.Trim().Equals("text/css", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        return string.IsNullOrWhiteSpace(style.Attribute("title")?.Value);
    }

    private static bool ParseSvgCssRules(
        string css,
        ICollection<SvgCssRule> rules,
        ref int declarationCount,
        ref int unsupported) {
        if (ContainsNonSvgCssWhitespace(css)) {
            unsupported++;
            return false;
        }
        int cursor = 0;
        while (cursor < css.Length && rules.Count < MaximumSvgCssRules && declarationCount < MaximumSvgCssDeclarations) {
            int open = FindSvgCssCharacter(css, '{', cursor);
            if (open < 0) break;
            int close = FindSvgCssBlockEnd(css, open + 1);
            if (close < 0) {
                unsupported++;
                return false;
            }
            string selectorText = css.Substring(cursor, open - cursor).Trim();
            string body = css.Substring(open + 1, close - open - 1);
            if (selectorText.StartsWith("@", StringComparison.Ordinal)) {
                unsupported++;
                return false;
            } else {
                IReadOnlyList<SvgCssDeclaration> declarations = ParseSvgCssDeclarations(body, ref declarationCount, out bool declarationsComplete);
                if (!declarationsComplete) {
                    unsupported++;
                    return false;
                }
                IReadOnlyList<string> selectors = SplitSvgCssTopLevel(selectorText, ',');
                foreach (string selector in selectors) {
                    string normalized = selector.Trim();
                    if (normalized.Length == 0) {
                        unsupported++;
                        return false;
                    }
                    if (declarations.Count == 0) continue;
                    if (rules.Count >= MaximumSvgCssRules) {
                        unsupported++;
                        return false;
                    }
                    if (!TryCalculateSvgSpecificity(
                            normalized,
                            out SvgCssSpecificity specificity,
                            out IReadOnlyList<SvgSelectorPart> parts)) {
                        unsupported++;
                        return false;
                    }
                    rules.Add(new SvgCssRule(normalized, parts, declarations, specificity, rules.Count));
                }
            }
            cursor = close + 1;
        }
        string remainder = css.Substring(cursor).Trim();
        if (remainder.Length > 0) {
            unsupported++;
            return false;
        }
        if (cursor < css.Length && (rules.Count >= MaximumSvgCssRules || declarationCount >= MaximumSvgCssDeclarations)) {
            unsupported++;
            return false;
        }
        return true;
    }

    private static IReadOnlyList<SvgCssDeclaration> ParseSvgCssDeclarations(string? text) {
        int ignored = 0;
        return ParseSvgCssDeclarations(text, ref ignored, out _);
    }

    private static IReadOnlyList<SvgCssDeclaration> ParseSvgCssDeclarations(
        string? text,
        ref int declarationCount,
        out bool complete) {
        var result = new List<SvgCssDeclaration>();
        complete = true;
        if (string.IsNullOrWhiteSpace(text)) return result;
        if (ContainsNonSvgCssWhitespace(text!)) {
            complete = false;
            return result;
        }
        foreach (string raw in SplitSvgCssTopLevel(text!, ';')) {
            if (declarationCount >= MaximumSvgCssDeclarations) {
                if (!string.IsNullOrWhiteSpace(raw)) complete = false;
                continue;
            }
            int colon = raw.IndexOf(':');
            if (colon <= 0) continue;
            string name = raw.Substring(0, colon).Trim();
            string value = raw.Substring(colon + 1).Trim();
            if (name.Length == 0 || value.Length == 0) continue;
            bool important = TryStripImportant(value, out value);
            result.Add(new SvgCssDeclaration(name, value, important));
            declarationCount++;
        }
        return result;
    }

    private static bool IsSvgPresentationPropertyName(string propertyName) => propertyName.ToLowerInvariant() switch {
        "baseline-shift" or "clip-path" or "clip-rule" or "color" or "display" or "dominant-baseline" or "fill" or "fill-opacity" or "fill-rule" or
        "filter" or "flood-color" or "flood-opacity" or "font-family" or "font-size" or "font-style" or "font-weight" or "line-height" or "marker-end" or
        "marker-mid" or "marker-start" or "mask" or "mask-type" or "mix-blend-mode" or "opacity" or "stop-color" or "stop-opacity" or
        "stroke" or "stroke-dasharray" or "stroke-dashoffset" or "stroke-linecap" or "stroke-linejoin" or
        "stroke-miterlimit" or "stroke-opacity" or "stroke-width" or "text-anchor" or "text-orientation" or "transform" or
        "vector-effect" or "visibility" or "writing-mode" => true,
        _ => false
    };

    private static bool IsSupportedSvgCssDeclaration(
        SvgCssDeclaration declaration,
        SvgPaintServerRegistry paintServers,
        ref int unsupported) {
        if (declaration.Name.StartsWith("--", StringComparison.Ordinal)) return true;
        int before = unsupported;
        string name = declaration.Name.Trim().ToLowerInvariant();
        string value = declaration.Value.Trim();
        if (!IsSvgPresentationPropertyName(name)) {
            return false;
        }
        if (IsSvgCssWideKeyword(value)) {
            if (value.Equals("revert", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("revert-layer", StringComparison.OrdinalIgnoreCase)) unsupported++;
            return unsupported == before;
        }
        if (value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0) return true;
        switch (name) {
            case "transform":
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                    !OfficeSvgTransformParser.TryParse(value, out _)) unsupported++;
                break;
            case "clip-path":
            case "filter":
            case "mask":
            case "marker-start":
            case "marker-mid":
            case "marker-end":
                if (!IsSvgLocalReferenceOrNone(value)) unsupported++;
                break;
            case "clip-rule":
                if (!value.Equals("nonzero", StringComparison.OrdinalIgnoreCase) &&
                    !value.Equals("evenodd", StringComparison.OrdinalIgnoreCase)) unsupported++;
                break;
            case "mix-blend-mode":
                if (!TryParseBlendMode(value, out _)) unsupported++;
                break;
            case "flood-color":
            case "stop-color":
                if (!value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase) &&
                    !OfficeColor.TryParseCss(value, out _)) unsupported++;
                break;
            case "flood-opacity":
            case "stop-opacity":
            case "opacity":
            case "fill-opacity":
            case "stroke-opacity":
                if (!TrySvgCssUnitOrPercentage(value)) unsupported++;
                break;
            case "mask-type":
                if (!value.Equals("alpha", StringComparison.OrdinalIgnoreCase) &&
                    !value.Equals("luminance", StringComparison.OrdinalIgnoreCase)) unsupported++;
                break;
            case "vector-effect":
                if (!value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                    !value.Equals("non-scaling-stroke", StringComparison.OrdinalIgnoreCase)) unsupported++;
                break;
            default:
                SvgPaintContext validation = SvgPaintContext.Default;
                ApplyProperty(name, value, paintServers, ref validation, ref unsupported);
                break;
        }
        return unsupported == before;
    }

    private static bool TryResolveSvgComputedCssValue(
        string propertyName,
        string authoredValue,
        IReadOnlyDictionary<string, string> customProperties,
        XElement? parent,
        SvgPaintServerRegistry paintServers,
        ref int unsupported,
        out string? computedValue) {
        string value;
        bool variablesResolved = TryResolveSvgCssVariables(
            authoredValue,
            customProperties,
            0,
            out value,
            out bool variableLimitExceeded);
        if (variableLimitExceeded) {
            computedValue = null;
            return false;
        }
        if (!variablesResolved) {
            value = "unset";
        } else if (!IsSvgCssWideKeyword(value)) {
            value = NormalizeSvgCssPercentageValue(propertyName, value);
            int validationUnsupported = 0;
            if (!IsSupportedSvgCssDeclaration(
                    new SvgCssDeclaration(propertyName, value, important: false),
                    paintServers,
                    ref validationUnsupported)) {
                unsupported += validationUnsupported;
                value = "unset";
            }
        }

        string normalized = value.Trim();
        if (normalized.Equals("revert", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("revert-layer", StringComparison.OrdinalIgnoreCase)) {
            computedValue = null;
            return false;
        }
        if ((normalized.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
             normalized.Equals("unset", StringComparison.OrdinalIgnoreCase) && IsInheritedSvgCssPropertyName(propertyName)) &&
            propertyName.Equals("baseline-shift", StringComparison.OrdinalIgnoreCase)) {
            computedValue = "inherit";
            return true;
        }
        if (normalized.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("unset", StringComparison.OrdinalIgnoreCase) && IsInheritedSvgCssPropertyName(propertyName)) {
            if (IsInheritedSvgCssPropertyName(propertyName)) {
                computedValue = null;
                return true;
            }
            computedValue = ResolveParentSvgCssValue(parent, propertyName);
            if (computedValue == null) return TryGetInitialSvgCssValue(propertyName, out computedValue);
            return true;
        }
        if (normalized.Equals("initial", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("unset", StringComparison.OrdinalIgnoreCase)) {
            return TryGetInitialSvgCssValue(propertyName, out computedValue);
        }
        computedValue = value;
        return true;
    }

    private static string NormalizeSvgCssPercentageValue(string propertyName, string value) {
        string name = propertyName.ToLowerInvariant();
        if (name is not "opacity" and not "fill-opacity" and not "stroke-opacity" and not "stop-opacity" and not "flood-opacity") {
            return value;
        }
        string normalized = value.Trim();
        if (!normalized.EndsWith("%", StringComparison.Ordinal) ||
            !double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float,
                CultureInfo.InvariantCulture, out double percentage) ||
            double.IsNaN(percentage) || double.IsInfinity(percentage)) return value;
        return Math.Max(0D, Math.Min(1D, percentage / 100D)).ToString("R", CultureInfo.InvariantCulture);
    }

    private static string? ResolveParentSvgCssValue(XElement? parent, string propertyName) {
        if (parent == null) return null;
        return ReadPresentationProperty(parent, propertyName)?.Trim();
    }

    private static bool IsSvgCssWideKeyword(string value) {
        string normalized = value.Trim();
        return normalized.Equals("initial", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("unset", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("revert", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("revert-layer", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUnsupportedSvgCssWideKeyword(string value) {
        string normalized = value.Trim();
        return normalized.Equals("revert", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("revert-layer", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsInheritedSvgCssPropertyName(string propertyName) => propertyName.ToLowerInvariant() switch {
        "color" or "fill" or "fill-opacity" or "fill-rule" or "font-family" or "font-size" or
        "font-style" or "font-weight" or "line-height" or "marker-end" or "marker-mid" or "marker-start" or
        "stroke" or "stroke-dasharray" or "stroke-dashoffset" or "stroke-linecap" or "stroke-linejoin" or
        "stroke-miterlimit" or "stroke-opacity" or "stroke-width" or "text-anchor" or "text-orientation" or
        "visibility" or "writing-mode" => true,
        _ => false
    };

    private static bool TryGetInitialSvgCssValue(string propertyName, out string? value) {
        value = propertyName.ToLowerInvariant() switch {
            "baseline-shift" => "0",
            "clip-path" => "none",
            "clip-rule" => "nonzero",
            "color" => "black",
            "display" => "inline",
            "dominant-baseline" => "auto",
            "fill" => "black",
            "fill-opacity" => "1",
            "fill-rule" => "nonzero",
            "filter" => "none",
            "flood-color" => "black",
            "flood-opacity" => "1",
            "font-family" => "Arial",
            "font-size" => "16",
            "font-style" => "normal",
            "font-weight" => "normal",
            "line-height" => "normal",
            "marker-end" or "marker-mid" or "marker-start" or "mask" => "none",
            "mask-type" => "luminance",
            "mix-blend-mode" => "normal",
            "opacity" => "1",
            "stop-color" => "black",
            "stop-opacity" => "1",
            "stroke" => "none",
            "stroke-dasharray" => "none",
            "stroke-dashoffset" => "0",
            "stroke-linecap" => "butt",
            "stroke-linejoin" => "miter",
            "stroke-miterlimit" => "4",
            "stroke-opacity" => "1",
            "stroke-width" => "1",
            "text-anchor" => "start",
            "text-orientation" => "mixed",
            "transform" => "none",
            "vector-effect" => "none",
            "visibility" => "visible",
            "writing-mode" => "horizontal-tb",
            _ => null
        };
        return value != null;
    }

    private static bool TrySvgCssUnitOrPercentage(string value) {
        if (TryUnit(value, out _)) return true;
        string normalized = value.Trim();
        if (!normalized.EndsWith("%", StringComparison.Ordinal)) return false;
        return double.TryParse(
            normalized.Substring(0, normalized.Length - 1),
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double percentage) && !double.IsNaN(percentage) && !double.IsInfinity(percentage);
    }

    private static bool IsSvgLocalReferenceOrNone(string value) {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        string normalized = value.Trim();
        if (!normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase) ||
            !normalized.EndsWith(")", StringComparison.Ordinal)) return false;
        string reference = normalized.Substring(4, normalized.Length - 5).Trim().Trim('\'', '"');
        return reference.Length > 1 && reference[0] == '#';
    }

    private static bool TryResolveSvgCssVariables(
        string value,
        IReadOnlyDictionary<string, string> customProperties,
        int depth,
        out string resolved,
        out bool limitExceeded) {
        resolved = value;
        limitExceeded = false;
        if (depth > 16 || resolved.Length > MaximumSvgComputedCssCharacters) {
            limitExceeded = true;
            return false;
        }
        int start = resolved.IndexOf("var(", StringComparison.OrdinalIgnoreCase);
        while (start >= 0) {
            int close = FindSvgCssBlockEnd(resolved, start + 4, '(', ')');
            if (close < 0) return false;
            string arguments = resolved.Substring(start + 4, close - start - 4);
            IReadOnlyList<string> parts = SplitSvgCssTopLevel(arguments, ',');
            string name = parts[0].Trim();
            if (!IsSupportedSvgCustomPropertyReferenceName(name)) return false;
            string replacement;
            if (!customProperties.TryGetValue(name, out replacement!)) {
                if (parts.Count < 2) return false;
                replacement = string.Join(",", parts.Skip(1)).Trim();
            }
            if (!TryResolveSvgCssVariables(
                    replacement,
                    customProperties,
                    depth + 1,
                    out replacement,
                    out bool nestedLimitExceeded)) {
                limitExceeded = nestedLimitExceeded;
                return false;
            }
            long expandedLength = (long)resolved.Length - (close - start + 1L) + replacement.Length;
            if (expandedLength > MaximumSvgComputedCssCharacters) {
                limitExceeded = true;
                return false;
            }
            resolved = resolved.Substring(0, start) + replacement + resolved.Substring(close + 1);
            start = resolved.IndexOf("var(", StringComparison.OrdinalIgnoreCase);
        }
        return true;
    }

    private static bool IsSupportedSvgCustomPropertyReferenceName(string name) {
        if (name.Length <= 2 || !name.StartsWith("--", StringComparison.Ordinal)) return false;
        for (int index = 2; index < name.Length; index++) {
            char character = name[index];
            if (char.IsHighSurrogate(character)) {
                if (index + 1 >= name.Length || !char.IsLowSurrogate(name[index + 1])) return false;
                index++;
                continue;
            }
            if (char.IsLowSurrogate(character) || char.IsControl(character) || char.IsWhiteSpace(character)) return false;
            if (character < 0x80 && !char.IsLetterOrDigit(character) && character != '-' && character != '_') return false;
        }
        return true;
    }

    private static IReadOnlyList<string> SplitSvgCssTopLevel(string text, char separator) {
        var result = new List<string>();
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && !IsEscapedSvgCssCharacter(text, index)) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') quote = current;
            else if (current is '(' or '[') depth++;
            else if (current is ')' or ']') depth--;
            else if (current == separator && depth == 0) {
                result.Add(text.Substring(start, index - start));
                start = index + 1;
            }
        }
        result.Add(text.Substring(start));
        return result;
    }

    private static int FindSvgCssCharacter(string text, char target, int start) {
        char quote = '\0';
        for (int index = start; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && !IsEscapedSvgCssCharacter(text, index)) quote = '\0';
            } else if (current is '\'' or '"') quote = current;
            else if (current == target) return index;
        }
        return -1;
    }

    private static int FindSvgCssBlockEnd(string text, int start, char open = '{', char close = '}') {
        int depth = 1;
        char quote = '\0';
        for (int index = start; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && !IsEscapedSvgCssCharacter(text, index)) quote = '\0';
            } else if (current is '\'' or '"') quote = current;
            else if (current == open) depth++;
            else if (current == close && --depth == 0) return index;
        }
        return -1;
    }

    private static string RemoveSvgCssComments(string css) {
        var result = new StringBuilder(css.Length);
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscapedSvgCssCharacter(css, index)) quote = '\0';
            } else if (current is '\'' or '"') {
                quote = current;
                result.Append(current);
            } else if (index + 1 < css.Length && current == '/' && css[index + 1] == '*') {
                int close = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (close < 0) break;
                index = close + 1;
            } else result.Append(current);
        }
        return result.ToString();
    }

    private static bool IsEscapedSvgCssCharacter(string text, int index) {
        int backslashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text[cursor] == '\\'; cursor--) backslashes++;
        return backslashes % 2 != 0;
    }

    private static bool ContainsNonSvgCssWhitespace(string text) {
        foreach (char character in text) {
            if (char.IsWhiteSpace(character) && !IsSvgCssWhitespace(character)) return true;
        }
        return false;
    }

    private readonly struct SvgCssDeclaration {
        internal SvgCssDeclaration(string name, string value, bool important) {
            string normalizedName = name.Trim();
            Name = normalizedName.StartsWith("--", StringComparison.Ordinal)
                ? normalizedName
                : normalizedName.ToLowerInvariant();
            Value = value;
            Important = important;
        }
        internal string Name { get; }
        internal string Value { get; }
        internal bool Important { get; }
    }

    private readonly struct SvgCssSpecificity : IComparable<SvgCssSpecificity> {
        internal static SvgCssSpecificity Inline => new SvgCssSpecificity(1, 0, 0, 0);

        internal SvgCssSpecificity(int inline, int ids, int classes, int types) {
            InlineCount = inline;
            IdCount = ids;
            ClassCount = classes;
            TypeCount = types;
        }

        internal int InlineCount { get; }
        internal int IdCount { get; }
        internal int ClassCount { get; }
        internal int TypeCount { get; }

        public int CompareTo(SvgCssSpecificity other) {
            int comparison = InlineCount.CompareTo(other.InlineCount);
            if (comparison != 0) return comparison;
            comparison = IdCount.CompareTo(other.IdCount);
            if (comparison != 0) return comparison;
            comparison = ClassCount.CompareTo(other.ClassCount);
            return comparison != 0 ? comparison : TypeCount.CompareTo(other.TypeCount);
        }
    }

    private readonly struct SvgCssWinner {
        internal SvgCssWinner(string value, bool important, SvgCssSpecificity specificity, int order) { Value = value; Important = important; Specificity = specificity; Order = order; }
        internal string Value { get; }
        internal bool Important { get; }
        internal SvgCssSpecificity Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct SvgCssRule {
        internal SvgCssRule(string selector, IReadOnlyList<SvgSelectorPart> parts, IReadOnlyList<SvgCssDeclaration> declarations, SvgCssSpecificity specificity, int order) { Selector = selector; Parts = parts; Declarations = declarations; Specificity = specificity; Order = order; }
        internal string Selector { get; }
        internal IReadOnlyList<SvgSelectorPart> Parts { get; }
        internal IReadOnlyList<SvgCssDeclaration> Declarations { get; }
        internal SvgCssSpecificity Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct SvgSelectorPart {
        internal SvgSelectorPart(string compound, bool directParent) { Compound = compound; DirectParent = directParent; }
        internal string Compound { get; }
        internal bool DirectParent { get; }
    }
}
