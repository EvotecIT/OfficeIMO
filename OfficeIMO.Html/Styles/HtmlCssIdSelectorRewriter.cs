using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>Source-preserving identifier edits for the bounded publication stylesheet profile.</summary>
internal static class HtmlCssIdSelectorRewriter {
    private static readonly HashSet<string> RelationshipAttributes = new HashSet<string>(new[] {
        "href", "name", "usemap", "for", "form", "list", "headers", "itemref", "aria-labelledby", "aria-describedby",
        "aria-controls", "aria-owns", "aria-flowto", "aria-activedescendant", "aria-details", "aria-errormessage",
        // These values may be rebased or have fragment targets repaired by publication resource rewriting.
        "src", "poster", "cite", "longdesc", "data", "definitionURL", "srcset", "imagesrcset", "style",
        "fill", "stroke", "filter", "clip-path", "mask", "marker", "marker-start", "marker-mid", "marker-end",
        "cursor", "ping", "archive"
    }, StringComparer.OrdinalIgnoreCase);
    internal static string Rewrite(string css, IReadOnlyDictionary<string, string> map, CancellationToken token,
        Func<string, string, string, HtmlCssAttributeSelectorEdit>? rewriteRelationship = null) {
        var edits = new List<(int Start, int Length, string Value)>();
        var introducedIds = new HashSet<string>(map.Where(pair => pair.Key != pair.Value).Select(pair => pair.Value), StringComparer.Ordinal);
        var blocks = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
        Rules(0, css.Length, 0, false);
        var result = new StringBuilder(css);
        foreach (var edit in edits.OrderByDescending(edit => edit.Start)) result.Remove(edit.Start, edit.Length).Insert(edit.Start, edit.Value);
        return result.ToString();

        void Rules(int start, int end, int depth, bool declarations) {
            if (depth > 64) throw new NotSupportedException("Selector reconciliation exceeds 64 rule levels.");
            int cursor = start;
            while (cursor < end) {
                token.ThrowIfCancellationRequested();
                Trivia(ref cursor, end);
                if (cursor == end) break;
                if (css[cursor] == ';') { cursor++; continue; }
                if (declarations && TryDeclaration(ref cursor, end)) continue;
                int prelude = cursor;
                string? at = null;
                if (css[cursor] == '@') {
                    cursor++;
                    if (!HtmlCssIdentifierParser.TryRead(css, ref cursor, out string name)) throw Unsupported();
                    at = name.ToLowerInvariant();
                }
                int delimiter = Delimiter(cursor, end);
                if (delimiter == end) throw Unsupported();
                if (css[delimiter] == ';') {
                    if (at != "charset" && at != "import" && at != "namespace" && at != "layer") throw Unsupported();
                    cursor = delimiter + 1;
                    continue;
                }
                if (!blocks.TryGetValue(delimiter, out int close) || close >= end) throw Unsupported();
                if (at == null) {
                    Selector(prelude, delimiter);
                    Rules(delimiter + 1, close, depth + 1, true);
                } else if (at == "media" || at == "supports" || at == "layer" || at == "container" || at == "scope") {
                    if (at == "scope") Selector(cursor, delimiter);
                    if (at == "supports") SupportSelectors(cursor, delimiter);
                    Rules(delimiter + 1, close, depth + 1, declarations);
                } else if (at != "font-face" && at != "page" && at != "counter-style" && at != "property" &&
                    at != "font-feature-values" && at != "keyframes" && at != "-webkit-keyframes") throw Unsupported();
                cursor = close + 1;
            }
        }

        bool TryDeclaration(ref int cursor, int end) {
            int valueStart = cursor;
            if (!HtmlCssIdentifierParser.TryRead(css, ref valueStart, out string property)) return false;
            Trivia(ref valueStart, end);
            if (valueStart >= end || css[valueStart] != ':') return false;
            valueStart++; Trivia(ref valueStart, end);
            bool custom = property.StartsWith("--", StringComparison.Ordinal);
            int delimiter = Delimiter(valueStart, end);
            if (delimiter < end && css[delimiter] == '{' && !custom && delimiter != valueStart)
                return false; // A type-selector pseudo-class, such as h2:hover { ... }, is a nested rule.
            int i = valueStart;
            while (i < end) {
                token.ThrowIfCancellationRequested();
                if (Component(ref i, end)) continue;
                if (css[i] == ';') { cursor = i + 1; return true; }
                if (css[i] == '(' || css[i] == '[' || css[i] == '{') {
                    i = Closing(i, end, css[i] == '(' ? ')' : css[i] == '[' ? ']' : '}') + 1;
                    continue;
                }
                i++;
            }
            cursor = end; return true;
        }

        void SupportSelectors(int start, int end) {
            for (int i = start; i < end;) {
                token.ThrowIfCancellationRequested();
                if ((css[i] != '\\') && SkipLiteral(ref i, end)) continue;
                int nameStart = i;
                if (HtmlCssIdentifierParser.TryRead(css, ref i, out string name)) {
                    if (name.Equals("selector", StringComparison.OrdinalIgnoreCase) && i < end && css[i] == '(') {
                        int close = Closing(i, end, ')');
                        Selector(i + 1, close); i = close + 1;
                    }
                } else i = nameStart + 1;
            }
        }

        void Selector(int start, int end) {
            for (int i = start; i < end;) {
                token.ThrowIfCancellationRequested();
                if (SkipLiteral(ref i, end)) continue;
                if (css[i] == '[') {
                    int close = Closing(i, end, ']');
                    Attribute(i + 1, close); i = close + 1; continue;
                }
                if (css[i++] != '#') continue;
                int nameStart = i;
                if (!HtmlCssIdentifierParser.TryRead(css, ref i, out string id) || i > end) throw Unsupported();
                if (map.TryGetValue(id, out string? replacement)) {
                    if (replacement != id) edits.Add((nameStart, i - nameStart, Identifier(replacement)));
                } else if (introducedIds.Contains(id)) {
                    // A destination-only ID did not match the source. Keep it impossible without
                    // changing the hash selector's specificity, including inside :not/:is/:has.
                    edits.Add((nameStart - 1, 0, ":where([id~=\"\"])"));
                }
            }
        }

        void Attribute(int start, int end) {
            int i = start; Trivia(ref i, end);
            if (!HtmlCssIdentifierParser.TryRead(css, ref i, out string name)) throw Unsupported();
            Trivia(ref i, end);
            if (i < end && css[i] == '|' && (i + 1 == end || css[i + 1] != '='))
                throw new NotSupportedException("Namespaced attribute selectors require explicit reconciliation.");
            bool relationship = RelationshipAttributes.Contains(name);
            if (!relationship && !name.Equals("id", StringComparison.OrdinalIgnoreCase)) return;
            if (relationship && rewriteRelationship == null)
                throw new NotSupportedException("Selectors on identifier relationships or rewritten resource attributes require explicit reconciliation.");
            if (i == end) {
                if (relationship) rewriteRelationship!(name, string.Empty, string.Empty);
                return; // Presence remains true after replacement.
            }
            string operation = "=";
            if ("~|^$*".IndexOf(css[i]) >= 0) { operation = css[i] + "="; i++; }
            if (i == end || css[i++] != '=') throw new NotSupportedException("Only supported attribute comparisons can be reconciled.");
            Trivia(ref i, end);
            int valueStart = i;
            string id;
            if (i < end && (css[i] == '\'' || css[i] == '"')) {
                char quote = css[i++]; int textStart = i;
                while (i < end && css[i] != quote) Advance(ref i);
                if (i == end) throw Unsupported();
                id = HtmlCssEscapeDecoder.Decode(css.Substring(textStart, i - textStart)); i++;
            } else if (!HtmlCssIdentifierParser.TryRead(css, ref i, out id)) throw Unsupported();
            int valueEnd = i; Trivia(ref i, end);
            if (i < end) {
                if (!HtmlCssIdentifierParser.TryRead(css, ref i, out string flag) || !flag.Equals("s", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Case-insensitive id selectors require explicit reconciliation.");
                Trivia(ref i, end);
            }
            if (i != end) throw Unsupported();
            if (!relationship && operation != "=" && rewriteRelationship == null)
                throw new NotSupportedException("Partial ID selectors require source attribute evidence.");
            HtmlCssAttributeSelectorEdit replacement = rewriteRelationship != null ? rewriteRelationship(name, operation, id) :
                HtmlCssAttributeSelectorEdit.Operand(map.TryGetValue(id, out string? mapped) ? mapped : id);
            if (replacement.ExactValues != null) {
                string expanded = ExpandAttribute(name, replacement.ExactValues);
                edits.Add((start - 1, end - start + 2, expanded));
            } else if (replacement.Value != id)
                edits.Add((valueStart, valueEnd - valueStart, HtmlCssStringEncoder.Quote(replacement.Value!)));
        }

        string ExpandAttribute(string name, IReadOnlyList<string> values) {
            if (values.Count > HtmlCssAttributeSelectorEdit.MaximumAlternatives) throw Unsupported();
            string attribute = Identifier(name);
            // Empty ~= never matches, but still contributes one attribute selector's specificity.
            if (values.Count == 0) return "[" + attribute + "~=\"\"]";
            var output = new StringBuilder();
            if (values.Count > 1) output.Append(":is(");
            foreach (string value in values) {
                token.ThrowIfCancellationRequested();
                if (value.Length > HtmlCssAttributeSelectorEdit.MaximumSelectorLength) throw Unsupported();
                if (output.Length > 4) output.Append(',');
                output.Append('[').Append(attribute).Append('=').Append(HtmlCssStringEncoder.Quote(value)).Append(']');
                if (output.Length > HtmlCssAttributeSelectorEdit.MaximumSelectorLength - 1) throw Unsupported();
            }
            if (values.Count > 1) output.Append(')');
            return output.ToString();
        }

        int Delimiter(int start, int end) {
            for (int i = start; i < end;) {
                token.ThrowIfCancellationRequested();
                if (Component(ref i, end)) continue;
                if (css[i] == '(' || css[i] == '[') { i = Closing(i, end, css[i] == '(' ? ')' : ']') + 1; continue; }
                if (css[i] == '{' || css[i] == ';') return i;
                if (css[i] == '}') throw Unsupported();
                i++;
            }
            return end;
        }

        int Closing(int start, int end, char close) {
            var endings = new Stack<char>(); endings.Push(close);
            for (int i = start + 1; i < end;) {
                token.ThrowIfCancellationRequested();
                if (Component(ref i, end)) continue;
                char current = css[i];
                if (current == '(' || current == '[' || current == '{') {
                    if (endings.Count >= 256) throw Unsupported();
                    endings.Push(current == '(' ? ')' : current == '[' ? ']' : '}');
                } else if (current == ')' || current == ']' || current == '}') {
                    if (current != endings.Pop()) throw Unsupported();
                    if (endings.Count == 0) return i;
                }
                i++;
            }
            throw Unsupported();
        }

        bool Component(ref int i, int end) {
            int nameEnd = i;
            if (HtmlCssIdentifierParser.TryRead(css, ref nameEnd, out string name)) {
                if (name.Equals("url", StringComparison.OrdinalIgnoreCase) && nameEnd < end && css[nameEnd] == '(') {
                    int value = nameEnd + 1;
                    while (value < end && char.IsWhiteSpace(css[value])) value++;
                    if (value < end && css[value] != '\'' && css[value] != '"') {
                        // An unquoted URL token owns slash/star and escaped delimiters as path data.
                        while (value < end && css[value] != ')') Advance(ref value);
                        if (value == end) throw Unsupported();
                        i = value + 1; return true;
                    }
                }
                i = nameEnd; return true;
            }
            return SkipLiteral(ref i, end);
        }

        bool SkipLiteral(ref int i, int end) {
            if (css[i] == '/' && i + 1 < end && css[i + 1] == '*') {
                int close = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                if (close < 0 || close + 2 > end) throw Unsupported();
                i = close + 2; return true;
            }
            if (css[i] == '\'' || css[i] == '"') {
                char quote = css[i++];
                while (i < end && css[i] != quote) Advance(ref i);
                if (i == end) throw Unsupported();
                i++; return true;
            }
            if (css[i] == '\\') { Advance(ref i); return true; }
            return false;
        }

        void Trivia(ref int i, int end) {
            while (i < end) {
                if (char.IsWhiteSpace(css[i])) i++;
                else if (css[i] == '/' && i + 1 < end && css[i + 1] == '*') SkipLiteral(ref i, end);
                else break;
            }
        }

        void Advance(ref int i) {
            if (css[i] == '\\' && HtmlCssEscapeDecoder.TryDecodeEscape(css, i, out _, out int consumed)) i += consumed;
            else i++;
        }
    }

    private static string Identifier(string value) {
        var result = new StringBuilder();
        foreach (char c in value) {
            if (char.IsLetterOrDigit(c) || c == '-' || c == '_' || c >= 128) result.Append(c);
            else result.Append('\\').Append(((int)c).ToString("x", CultureInfo.InvariantCulture)).Append(' ');
        }
        return result.ToString();
    }

    private static NotSupportedException Unsupported() => new NotSupportedException("Stylesheet syntax is outside the supported selector reconciliation profile.");
}
