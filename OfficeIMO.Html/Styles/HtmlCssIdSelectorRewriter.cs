using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>Source-preserving identifier edits for the bounded publication stylesheet profile.</summary>
internal static class HtmlCssIdSelectorRewriter {
    private static readonly HashSet<string> RelationshipAttributes = new HashSet<string>(new[] {
        "href", "name", "usemap", "for", "form", "list", "headers", "aria-labelledby", "aria-describedby",
        "aria-controls", "aria-owns", "aria-flowto", "aria-activedescendant", "aria-details", "aria-errormessage"
    }, StringComparer.OrdinalIgnoreCase);
    internal static string Rewrite(string css, IReadOnlyDictionary<string, string> map, CancellationToken token) {
        var edits = new List<(int Start, int Length, string Value)>();
        var blocks = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
        int[] openings = blocks.Keys.OrderBy(value => value).ToArray();
        Rules(0, css.Length, 0);
        var result = new StringBuilder(css);
        foreach (var edit in edits.OrderByDescending(edit => edit.Start)) result.Remove(edit.Start, edit.Length).Insert(edit.Start, edit.Value);
        return result.ToString();

        void Rules(int start, int end, int depth) {
            if (depth > 64) throw new NotSupportedException("Selector reconciliation exceeds 64 rule levels.");
            int cursor = start;
            while (cursor < end) {
                token.ThrowIfCancellationRequested();
                Trivia(ref cursor, end);
                if (cursor == end) break;
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
                    // Nested style rules and custom-property blocks need a separate grammar.
                    int next = Array.BinarySearch(openings, delimiter) + 1;
                    if (next < openings.Length && openings[next] < close) throw new NotSupportedException("Selector reconciliation does not support nested style declarations.");
                } else if (at == "media" || at == "supports" || at == "layer" || at == "container" || at == "scope") {
                    if (at == "scope") Selector(cursor, delimiter);
                    if (at == "supports") SupportSelectors(cursor, delimiter);
                    Rules(delimiter + 1, close, depth + 1);
                } else if (at != "font-face" && at != "page" && at != "counter-style" && at != "property" &&
                    at != "font-feature-values" && at != "keyframes" && at != "-webkit-keyframes") throw Unsupported();
                cursor = close + 1;
            }
        }

        void SupportSelectors(int start, int end) {
            for (int i = start; i < end;) {
                token.ThrowIfCancellationRequested();
                if ((css[i] != '\\') && SkipLiteral(ref i, end)) continue;
                int nameStart = i;
                if (HtmlCssIdentifierParser.TryRead(css, ref i, out string name)) {
                    if (name.Equals("selector", StringComparison.OrdinalIgnoreCase) && i < end && css[i] == '(') {
                        int close = Closing(i, end, '(', ')');
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
                    int close = Closing(i, end, '[', ']');
                    Attribute(i + 1, close); i = close + 1; continue;
                }
                if (css[i++] != '#') continue;
                int nameStart = i;
                if (!HtmlCssIdentifierParser.TryRead(css, ref i, out string id) || i > end) throw Unsupported();
                if (map.TryGetValue(id, out string? replacement) && replacement != id)
                    edits.Add((nameStart, i - nameStart, Identifier(replacement)));
            }
        }

        void Attribute(int start, int end) {
            int i = start; Trivia(ref i, end);
            if (!HtmlCssIdentifierParser.TryRead(css, ref i, out string name)) throw Unsupported();
            Trivia(ref i, end);
            if (i < end && css[i] == '|' && (i + 1 == end || css[i + 1] != '='))
                throw new NotSupportedException("Namespaced attribute selectors require explicit reconciliation.");
            if (RelationshipAttributes.Contains(name))
                throw new NotSupportedException("Selectors on identifier relationships require explicit reconciliation.");
            if (!name.Equals("id", StringComparison.OrdinalIgnoreCase)) return;
            if (i == end) return; // Presence remains true after replacement.
            if (css[i++] != '=') throw new NotSupportedException("Only exact id attribute selectors can be reconciled.");
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
            if (map.TryGetValue(id, out string? replacement) && replacement != id)
                edits.Add((valueStart, valueEnd - valueStart, HtmlCssStringEncoder.Quote(replacement)));
        }

        int Delimiter(int start, int end) {
            for (int i = start; i < end;) {
                token.ThrowIfCancellationRequested();
                if (SkipLiteral(ref i, end)) continue;
                if (css[i] == '(' || css[i] == '[') { i = Closing(i, end, css[i], css[i] == '(' ? ')' : ']') + 1; continue; }
                if (css[i] == '{' || css[i] == ';') return i;
                if (css[i] == '}') throw Unsupported();
                i++;
            }
            return end;
        }

        int Closing(int start, int end, char open, char close) {
            int depth = 1;
            for (int i = start + 1; i < end;) {
                token.ThrowIfCancellationRequested();
                if (SkipLiteral(ref i, end)) continue;
                if (css[i] == open) depth++;
                if (css[i] == close && --depth == 0) return i;
                i++;
            }
            throw Unsupported();
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
