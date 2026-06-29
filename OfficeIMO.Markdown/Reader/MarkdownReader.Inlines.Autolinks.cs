namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsAngleAutolinkStart(string text, int start) {
        if (start < 0 || start >= text.Length) return false;
        if (text[start] != '<') return false;
        return TryParseAngleAutolink(text, start, out _, out _, out _);
    }

    private static bool TryParseAngleAutolink(string text, int start, out int consumed, out string label, out string href) {
        consumed = 0;
        label = href = string.Empty;
        if (start < 0 || start >= text.Length || text[start] != '<') return false;
        int gt = text.IndexOf('>', start + 1);
        if (gt < 0) return false;
        if (gt == start + 1) return false;

        // Disallow whitespace/control inside.
        for (int i = start + 1; i < gt; i++) {
            char c = text[i];
            if (char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }

        var inner = text.Substring(start + 1, gt - (start + 1));

        // URL form
        if (inner.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            inner.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        if (inner.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        if (TryGetScheme(inner, out var scheme) && IsUriAngleAutolink(inner, scheme)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        // Email form
        if (LooksLikeEmail(inner)) {
            label = inner;
            href = "mailto:" + inner;
            consumed = gt - start + 1;
            return true;
        }

        return false;
    }

    private static bool IsUriAngleAutolink(string inner, string scheme) {
        if (string.IsNullOrEmpty(inner) || string.IsNullOrEmpty(scheme)) return false;

        // Match CommonMark-style absolute URI autolinks instead of limiting support to scheme://...
        // This keeps tel:, urn:, xmpp:, etc. on the same policy-controlled path as http(s)/mailto.
        if (scheme.Length < 2 || scheme.Length > 32) return false;
        if (inner.Length <= scheme.Length + 1) return false;

        for (int i = scheme.Length + 1; i < inner.Length; i++) {
            char c = inner[i];
            if (char.IsWhiteSpace(c) || char.IsControl(c) || c == '<' || c == '>') return false;
        }

        return true;
    }

    private static bool LooksLikeEmail(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        int at = s.IndexOf('@');
        if (at <= 0 || at == s.Length - 1) return false;
        // Reject "mailto:" which is a URL form and will be handled above if ever enabled.
        if (s.IndexOf(':') >= 0) return false;

        string domain = s.Substring(at + 1);
        // Require at least one '.' in domain and not at the ends.
        int dot = domain.IndexOf('.');
        if (dot <= 0 || dot == domain.Length - 1) return false;
        if (!char.IsLetterOrDigit(domain[domain.Length - 1])) return false;

        // Basic character checks (no spaces/control already enforced by caller).
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '@') continue;
            if (i > at && c == '+') return false;
            if (c == '.' || c == '-' || c == '_' || c == '+') continue;
            if (char.IsLetterOrDigit(c)) continue;
            return false;
        }
        return true;
    }

    private static bool StartsWithHttp(string text, int start, MarkdownReaderOptions options, out int end) {
        end = start;
        if (start + 7 > text.Length) return false;
        // Require a boundary on the left so we don't linkify inside longer words.
        if (HasInvalidAutolinkLeftBoundary(text, start, options)) return false;
        if (IsAfterInvalidReferenceDefinitionPrefix(text, start)) return false;
        var rem = text.Substring(start);
        if (!(rem.StartsWith("http://") || rem.StartsWith("https://"))) return false;
        int rawEnd = ConsumeLiteralUrl(text, start, options);
        int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
        if (ShouldRejectUnmatchedOpeningSingleQuote(text, start, rawEnd, i)) return false;
        if (options.AutolinkRejectUserInfoAuthority && AutolinkAuthorityContainsUserInfo(text, start, i)) return false;
        if (!options.AutolinkAllowQueryAndFragmentSpecialCharacters && ShouldRejectQueryFragmentSpecialCharsAutolink(text, start, i)) return false;
        if (!options.AutolinkAllowBalancedParenthesesWithTrailingPunctuation && ShouldRejectAmbiguousTrailingParen(text, start, rawEnd, i)) return false;
        if (!options.AutolinkAllowDomainWithoutPeriod && !HttpAutolinkHasDomainPeriod(text, start, i)) return false;
        end = i; return end > start + 7;
    }

    private static bool StartsWithWww(string text, int start, MarkdownReaderOptions options, out int end) {
        end = start;
        if (start + 4 > text.Length) return false;
        if (HasInvalidAutolinkLeftBoundary(text, start, options)) return false;
        if (IsAfterInvalidReferenceDefinitionPrefix(text, start)) return false;
        if (options.AutolinkRequireLowercaseWwwPrefix) {
            if (!text.Substring(start).StartsWith("www.", StringComparison.Ordinal)) return false;
        } else if (!text.Substring(start).StartsWith("www.", StringComparison.OrdinalIgnoreCase)) return false;

        int rawEnd = ConsumeLiteralUrl(text, start, options);
        int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
        int scanEnd = rawEnd;
        if (ShouldRejectUnmatchedOpeningSingleQuote(text, start, rawEnd, i)) return false;
        if (options.AutolinkRejectUserInfoAuthority && AutolinkAuthorityContainsUserInfo(text, start, i)) return false;
        if (!options.AutolinkAllowQueryAndFragmentSpecialCharacters && ShouldRejectQueryFragmentSpecialCharsAutolink(text, start, i)) return false;
        if (!options.AutolinkAllowBalancedParenthesesWithTrailingPunctuation && ShouldRejectAmbiguousTrailingParen(text, start, rawEnd, i)) return false;

        // Must include at least one dot after the www.
        var token = text.Substring(start, i - start);
        if (token.Length <= 4) return false;
        if (!options.AutolinkAllowDomainWithoutPeriod && token.IndexOf('.', 4) < 0) return false;
        if (!IsGfmWwwHostAllowed(token, options.AutolinkAllowDomainWithoutPeriod, options.AutolinkRejectUnderscoreInWwwHost)) return false;

        // Right boundary: avoid linking as part of an identifier-like token.
        if (scanEnd < text.Length && IsEmailChar(text[scanEnd])) return false;

        end = i;
        return end > start + 4;
    }

    private static bool IsGfmWwwHostAllowed(string token, bool allowDomainWithoutPeriod, bool rejectUnderscoreInHost) {
        if (string.IsNullOrEmpty(token) || !token.StartsWith("www.", StringComparison.OrdinalIgnoreCase)) return false;

        int hostEnd = token.Length;
        for (int i = 4; i < token.Length; i++) {
            char c = token[i];
            if (c == '/' || c == '?' || c == '#') {
                hostEnd = i;
                break;
            }
        }

        if (hostEnd <= 4) return false;

        string host = token.Substring(0, hostEnd);
        string[] labels = host.Split('.');
        if (labels.Length < (allowDomainWithoutPeriod ? 2 : 3)) return false;

        for (int i = 1; i < labels.Length; i++) {
            string label = labels[i];
            if (label.Length == 0) {
                return false;
            }

            if (rejectUnderscoreInHost && label.IndexOf('_') >= 0) {
                return false;
            }

            if (!rejectUnderscoreInHost && i >= 2 && (label[0] == '_' || label.IndexOf('_') >= 0)) {
                return false;
            }
        }

        return true;
    }

    private static bool AutolinkAuthorityContainsUserInfo(string text, int start, int end) {
        if (string.IsNullOrEmpty(text) || start < 0 || end <= start || end > text.Length) return false;

        int authorityStart = start;
        int schemeSeparator = text.IndexOf("://", start, end - start, StringComparison.Ordinal);
        if (schemeSeparator >= start) {
            authorityStart = schemeSeparator + 3;
        }

        int authorityEnd = end;
        for (int i = authorityStart; i < end; i++) {
            char c = text[i];
            if (c == '/' || c == '?' || c == '#') {
                authorityEnd = i;
                break;
            }
        }

        return text.IndexOf('@', authorityStart, authorityEnd - authorityStart) >= 0;
    }

    private static bool TryConsumeBareSchemeAutolink(string text, int start, MarkdownReaderOptions options, out int end, out string label, out string href) {
        end = start;
        label = href = string.Empty;
        if (start < 0 || start >= text.Length) return false;
        if (HasInvalidAutolinkLeftBoundary(text, start, options)) return false;
        if (IsAfterInvalidReferenceDefinitionPrefix(text, start)) return false;

        if (IsBareSchemePrefixEnabled(options, "mailto:") &&
            StartsWithAutolinkScheme(text, start, "mailto:", options)) {
            int emailStart = start + "mailto:".Length;
            if (!TryConsumeBareMailtoAddress(text, emailStart, options, out int emailEnd, out string email)) return false;
            end = emailEnd;
            label = options.AutolinkBareMailtoDisplayAddressOnly
                ? email
                : text.Substring(start, end - start);
            href = "mailto:" + email;
            return true;
        }

        if (IsBareSchemePrefixEnabled(options, "ftp://") &&
            StartsWithAutolinkScheme(text, start, "ftp://", options)) {
            int rawEnd = ConsumeLiteralUrl(text, start, options);
            int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
            if (ShouldRejectUnmatchedOpeningSingleQuote(text, start, rawEnd, i)) return false;
            if (options.AutolinkRejectUserInfoAuthority && AutolinkAuthorityContainsUserInfo(text, start, i)) return false;
            if (!options.AutolinkAllowQueryAndFragmentSpecialCharacters && ShouldRejectQueryFragmentSpecialCharsAutolink(text, start, i)) return false;
            if (!options.AutolinkAllowBalancedParenthesesWithTrailingPunctuation && ShouldRejectAmbiguousTrailingParen(text, start, rawEnd, i)) return false;
            if (!options.AutolinkAllowDomainWithoutPeriod && !HttpAutolinkHasDomainPeriod(text, start, i)) return false;
            if (i <= start + "ftp://".Length) return false;
            end = i;
            label = text.Substring(start, end - start);
            href = label;
            return true;
        }

        if (IsBareSchemePrefixEnabled(options, "tel:") &&
            StartsWithAutolinkScheme(text, start, "tel:", options)) {
            int valueStart = start + "tel:".Length;
            int rawEnd = ConsumeLiteralUrl(text, start, options);
            int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
            if (i <= valueStart) return false;
            end = i;
            label = text.Substring(valueStart, end - valueStart);
            href = text.Substring(start, end - start);
            return true;
        }

        if (IsBareSchemePrefixEnabled(options, "xmpp:") &&
            StartsWithAutolinkScheme(text, start, "xmpp:", options)) {
            int rawEnd = ConsumeLiteralUrl(text, start, options);
            int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
            if (i <= start + "xmpp:".Length) return false;
            end = i;
            label = text.Substring(start, end - start);
            href = label;
            return true;
        }

        return false;
    }

    private static bool IsBareSchemePrefixEnabled(MarkdownReaderOptions options, string prefix) {
        if (options.AutolinkBareSchemePrefixes == null) {
            return true;
        }

        for (int i = 0; i < options.AutolinkBareSchemePrefixes.Length; i++) {
            if (string.Equals(options.AutolinkBareSchemePrefixes[i], prefix, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool StartsWithAutolinkScheme(string text, int start, string scheme, MarkdownReaderOptions options) {
        if (options.AutolinkRequireLowercaseBareSchemePrefix) {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(scheme)) return false;
            if (start < 0 || start + scheme.Length > text.Length) return false;
            return string.Compare(text, start, scheme, 0, scheme.Length, StringComparison.Ordinal) == 0;
        }

        return StartsWithOrdinalIgnoreCase(text, start, scheme);
    }

    private static bool TryConsumeBareMailtoAddress(string text, int start, MarkdownReaderOptions options, out int end, out string email) {
        end = start;
        email = string.Empty;
        if (start < 0 || start >= text.Length) return false;

        int rawEnd = ConsumeLiteralUrl(text, start, options);
        int i = TrimTrailingAutolinkPunctuation(text, start, rawEnd, options);
        if (i <= start) return false;

        int addressEnd = i;
        for (int scan = start; scan < i; scan++) {
            char c = text[scan];
            if (c == '/' || c == '?' || c == '#') {
                addressEnd = scan;
                break;
            }
        }

        if (addressEnd <= start) return false;

        string address = text.Substring(start, addressEnd - start);
        if (!LooksLikeEmail(address)) return false;

        end = i;
        email = text.Substring(start, i - start);
        return true;
    }

    private static bool StartsWithOrdinalIgnoreCase(string text, int start, string value) {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value)) return false;
        if (start < 0 || start + value.Length > text.Length) return false;
        return string.Compare(text, start, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsBareSchemeAutolinkStartCandidate(char c) {
        return c == 'f' || c == 'F'
            || c == 'm' || c == 'M'
            || c == 't' || c == 'T'
            || c == 'x' || c == 'X';
    }

    private static int TrimTrailingAutolinkPunctuation(string text, int start, int rawEnd, MarkdownReaderOptions options) {
        if (options.AutolinkTrimSingleTrailingPunctuationOrUnderscore) {
            return TrimSingleTrailingAutolinkPunctuationOrUnderscore(text, start, rawEnd, options);
        }

        int i = rawEnd;
        bool removedClosingParenthesis = false;
        bool changed;
        do {
            changed = false;

            int parenTrimmed = TrimUnmatchedTrailingClosingParentheses(text, start, i);
            if (parenTrimmed != i) {
                removedClosingParenthesis = true;
                i = parenTrimmed;
                changed = true;
            }

            int entitySuffixStart = FindTrailingEntityLikeSuffixStart(text, start, i);
            if (entitySuffixStart >= 0) {
                i = entitySuffixStart;
                changed = true;
                continue;
            }

            while (i > start && IsTrailingAutolinkPunctuation(text[i - 1])) {
                if (options.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis && removedClosingParenthesis) {
                    break;
                }

                i--;
                changed = true;
            }
        } while (changed && i > start);

        return i;
    }

    private static int TrimSingleTrailingAutolinkPunctuationOrUnderscore(string text, int start, int rawEnd, MarkdownReaderOptions options) {
        int i = rawEnd;
        bool removedClosingParenthesis = false;
        bool trimmedSingleDelimiter = false;

        while (i > start) {
            int parenTrimmed = TrimUnmatchedTrailingClosingParentheses(text, start, i);
            if (parenTrimmed != i) {
                removedClosingParenthesis = true;
                i = parenTrimmed;
                continue;
            }

            int entitySuffixStart = FindTrailingEntityLikeSuffixStart(text, start, i);
            if (entitySuffixStart >= 0) {
                i = entitySuffixStart;
                continue;
            }

            char last = text[i - 1];
            if (!trimmedSingleDelimiter && IsTrailingAutolinkPunctuationOrUnderscore(last)) {
                if (options.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis
                    && removedClosingParenthesis
                    && IsTrailingAutolinkPunctuation(last)) {
                    break;
                }

                i--;
                trimmedSingleDelimiter = true;
                continue;
            }

            break;
        }

        return i;
    }

    private static int TrimUnmatchedTrailingClosingParentheses(string text, int start, int end) {
        int i = end;
        while (i > start && text[i - 1] == ')' && HasMoreClosingThanOpeningParentheses(text, start, i)) {
            i--;
        }

        return i;
    }

    private static bool HasMoreClosingThanOpeningParentheses(string text, int start, int end) {
        int balance = 0;
        for (int i = start; i < end; i++) {
            if (text[i] == '(') {
                balance++;
            } else if (text[i] == ')') {
                balance--;
            }
        }

        return balance < 0;
    }

    private static int FindTrailingEntityLikeSuffixStart(string text, int start, int end) {
        if (end - start < 3 || text[end - 1] != ';') {
            return -1;
        }

        int ampersand = text.LastIndexOf('&', end - 2, end - start - 1);
        if (ampersand < start || ampersand + 1 >= end - 1) {
            return -1;
        }

        for (int i = ampersand + 1; i < end - 1; i++) {
            if (!char.IsLetterOrDigit(text[i])) {
                return -1;
            }
        }

        return ampersand;
    }

    private static bool IsTrailingAutolinkPunctuation(char c) {
        return c == '.'
            || c == ','
            || c == ';'
            || c == ':'
            || c == '!'
            || c == '?'
            || c == '\''
            || c == '"';
    }

    private static bool IsTrailingAutolinkPunctuationOrUnderscore(char c) {
        return c == '_' || IsTrailingAutolinkPunctuation(c);
    }

    private static bool HasInvalidAutolinkLeftBoundary(string text, int start, MarkdownReaderOptions? options = null) {
        if (string.IsNullOrEmpty(text) || start <= 0 || start > text.Length) return false;

        char previous = text[start - 1];
        if (options?.AutolinkValidPreviousCharacters != null) {
            return !char.IsWhiteSpace(previous)
                && options.AutolinkValidPreviousCharacters.IndexOf(previous) < 0;
        }

        return char.IsLetterOrDigit(previous)
            || previous == '_'
            || previous == '/'
            || previous == ':'
            || previous == '.'
            || previous == '+'
            || previous == '-'
            || previous == '='
            || previous == '&'
            || (previous == '(' && options?.AutolinkAllowBalancedParenthesesWithTrailingPunctuation != true)
            || previous == '[';
    }

    private static bool ShouldRejectUnmatchedOpeningSingleQuote(string text, int start, int rawEnd, int trimmedEnd) {
        if (string.IsNullOrEmpty(text) || start <= 0) return false;
        if (text[start - 1] != '\'') return false;

        for (int i = trimmedEnd; i < rawEnd; i++) {
            if (text[i] == '\'') return false;
        }

        return true;
    }

    private static bool IsAfterInvalidReferenceDefinitionPrefix(string text, int start) {
        if (string.IsNullOrEmpty(text) || start <= 0 || start > text.Length) return false;

        int lineStart = text.LastIndexOf('\n', start - 1);
        lineStart = lineStart < 0 ? 0 : lineStart + 1;
        int lineEnd = text.IndexOf('\n', start);
        if (lineEnd < 0) lineEnd = text.Length;

        string line = text.Substring(lineStart, lineEnd - lineStart);
        if (!StartsWithReferenceDefinitionLikeLabel(line)) return false;

        return !TryParseReferenceLinkDefinition(new[] { line }, 0, new MarkdownReaderOptions(), out _, out _, out _, out _);
    }

    private static int ConsumeLiteralUrl(string text, int start, MarkdownReaderOptions options) {
        int i = start;
        int parenDepth = 0;
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ']' || c == '<') break;
            if (c == '(') {
                parenDepth++;
                i++;
                continue;
            }
            if (c == ')') {
                if (parenDepth == 0) {
                    if (!options.AutolinkAllowBalancedParenthesesWithTrailingPunctuation) break;
                } else {
                    parenDepth--;
                }
                i++;
                continue;
            }
            i++;
        }

        return i;
    }

    private static bool ShouldRejectAmbiguousTrailingParen(string text, int start, int rawEnd, int trimmedEnd) {
        if (string.IsNullOrEmpty(text) || start < 0 || trimmedEnd <= start) return false;

        bool extraClosingParenOutsideUrl = rawEnd < text.Length && text[rawEnd] == ')';
        bool trailingPunctuationTrimmedAfterBalancedParen = rawEnd > trimmedEnd && text[trimmedEnd - 1] == ')';
        if (!extraClosingParenOutsideUrl && !trailingPunctuationTrimmedAfterBalancedParen) return false;
        if (start > 0 && text[start - 1] == '(') return false;

        bool sawOpenParen = false;
        for (int i = start; i < trimmedEnd - 1; i++) {
            if (text[i] == '(') {
                sawOpenParen = true;
                break;
            }
        }

        return sawOpenParen;
    }

    private static bool ShouldRejectQueryFragmentSpecialCharsAutolink(string text, int start, int end) {
        if (string.IsNullOrEmpty(text) || start < 0 || end <= start) return false;

        int queryOrFragmentIndex = -1;
        for (int i = start; i < end; i++) {
            char ch = text[i];
            if (ch == '?' || ch == '#') {
                queryOrFragmentIndex = i;
                break;
            }
        }

        if (queryOrFragmentIndex < 0) return false;

        for (int i = queryOrFragmentIndex + 1; i < end; i++) {
            char ch = text[i];
            if (ch == '(' || ch == ')' || ch == '&') {
                return true;
            }
        }

        return false;
    }

    private static bool HttpAutolinkHasDomainPeriod(string text, int start, int end) {
        int schemeEnd = text.IndexOf("://", start, StringComparison.Ordinal);
        if (schemeEnd < 0 || schemeEnd + 3 >= end) return false;

        int hostStart = schemeEnd + 3;
        int hostEnd = end;
        for (int i = hostStart; i < end; i++) {
            char c = text[i];
            if (c == '/' || c == '?' || c == '#' || c == ':') {
                hostEnd = i;
                break;
            }
        }

        if (hostEnd <= hostStart) return false;
        if (text[hostStart] == '[') {
            return text.IndexOf(']', hostStart, hostEnd - hostStart) >= 0;
        }

        return text.IndexOf('.', hostStart, hostEnd - hostStart) >= 0;
    }

    private static bool TryConsumePlainEmail(string text, int start, MarkdownReaderOptions options, out int end, out string email) {
        end = start;
        email = string.Empty;
        if (start < 0 || start >= text.Length) return false;
        if (!IsEmailStartChar(text[start])) return false;
        if (options.AutolinkValidPreviousCharacters != null) {
            if (HasInvalidAutolinkLeftBoundary(text, start, options)) return false;
        } else if (start > 0 && (IsEmailChar(text[start - 1]) || text[start - 1] == '+' || text[start - 1] == '/' || text[start - 1] == ':' || text[start - 1] == '=' || text[start - 1] == '&' || text[start - 1] == '(' || text[start - 1] == '\'' || text[start - 1] == '[')) return false;
        if (IsImmediatelyAfterStandaloneMailtoScheme(text, start, options)) return false;

        int i = start;
        bool sawAt = false;
        // Stop at whitespace or common "outside token" delimiters; keep it pragmatic.
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ')' || c == ']' || c == '<') break;
            if (!IsEmailChar(c)) break;
            if (c == '@') sawAt = true;
            i++;
        }
        if (!sawAt) return false;

        int scanEnd = i;
        int j = i;
        while (j > start && (text[j - 1] == '.' || text[j - 1] == ',' || text[j - 1] == ';' || text[j - 1] == ':')) j--;
        if (j <= start) return false;

        var token = text.Substring(start, j - start);
        if (!LooksLikeEmail(token)) return false;
        if (scanEnd < text.Length) {
            if (IsEmailChar(text[scanEnd])) return false;
            if (text[scanEnd] == '/' || text[scanEnd] == '#') return false;
        }

        end = j;
        email = token;
        return true;
    }

    private static bool IsEmailStartChar(char c) => char.IsLetterOrDigit(c);

    private static bool IsEmailChar(char c) {
        if (char.IsLetterOrDigit(c)) return true;
        return c == '@' || c == '.' || c == '-' || c == '_' || c == '+';
    }

    private static bool IsImmediatelyAfterStandaloneMailtoScheme(string text, int start, MarkdownReaderOptions options) {
        if (string.IsNullOrEmpty(text) || start < 7) return false;
        if (text[start - 1] != ':') return false;

        int schemeStart = start - 7;
        return string.Compare(text, schemeStart, "mailto:", 0, 7, StringComparison.OrdinalIgnoreCase) == 0
            && !HasInvalidAutolinkLeftBoundary(text, schemeStart, options);
    }
}
