#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static bool IsPdfDelimiter(char value) {
        switch (value) {
            case '(':
            case ')':
            case '<':
            case '>':
            case '[':
            case ']':
            case '{':
            case '}':
            case '/':
            case '%':
                return true;
            default:
                return false;
        }
    }
    private static PdfObject? ParseTopLevelObject(
        string body,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) {
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (string.IsNullOrWhiteSpace(body)) return null;
        if (body.Length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, body.Length);
        }

        var s = body.TrimStart();
        if (string.Equals(s, "true", StringComparison.Ordinal)) return new PdfBoolean(true);
        if (string.Equals(s, "false", StringComparison.Ordinal)) return new PdfBoolean(false);
        if (string.Equals(s, "null", StringComparison.Ordinal)) return PdfNull.Instance;
        if (s.StartsWith("<<", System.StringComparison.Ordinal)) {
            // Find matching >> and parse inside
            int dictStart = body.IndexOf("<<", StringComparison.Ordinal);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(body, dictStart, body.Length);
                if (dictEnd > dictStart) {
                    if (SkipWhitespaceAndComments(body, dictEnd, body.Length) < body.Length) return null;
                    int dictionaryCharacters = dictEnd - (dictStart + 2);
                    if (dictionaryCharacters > effectiveLimits.MaxObjectCharacters) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, dictionaryCharacters);
                    }

                    try { return ParseDictionary(body, dictStart + 2, dictionaryCharacters, effectiveLimits, trackEncodedStringSourceSpans); } catch (PdfReadLimitException) { throw; } catch { return null; }
                }
            }
            return null;
        }
        if (s.Length > 0 && s[0] == '[') {
            using var toks = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, consumed) = ParseObject(toks, 0, effectiveLimits, 0);
            return consumed + 1 == toks.Count ? obj : null;
        }
        if (s.Length > 0 && s[0] == '(') {
            using var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, consumed) = ParseObject(stringTokens, 0, effectiveLimits, 0);
            return consumed + 1 == stringTokens.Count ? obj : null;
        }
        if (s.Length > 0 && s[0] == '<' && (s.Length == 1 || s[1] != '<')) {
            using var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, consumed) = ParseObject(stringTokens, 0, effectiveLimits, 0);
            return consumed + 1 == stringTokens.Count ? obj : null;
        }
        // number or name fallbacks
        using var tokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
        if (tokens.Count > 0) {
            var (obj0, consumed) = ParseObject(tokens, 0, effectiveLimits, 0);
            return consumed + 1 == tokens.Count ? obj0 : null;
        }
        return null;
    }

    internal static bool HasFlateDecode(PdfDictionary dict) {
        if (!dict.Items.TryGetValue("Filter", out var f)) return false;
        if (f is PdfName n) return string.Equals(n.Name, "FlateDecode", System.StringComparison.Ordinal);
        if (f is PdfArray arr) {
            foreach (var item in arr.Items) if (item is PdfName nn && string.Equals(nn.Name, "FlateDecode", System.StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static PdfDictionary ParseDictionary(
        string dict,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) =>
        ParseDictionary(dict, 0, dict.Length, limits, trackEncodedStringSourceSpans);

    private static PdfDictionary ParseDictionary(
        string source,
        int start,
        int length,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) {
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, length);
        }

        var d = new PdfDictionary();
        using var tokens = Tokenize(source, start, length, effectiveLimits, trackEncodedStringSourceSpans);
        for (int i = 0; i < tokens.Count; i++) {
            string tokenText = tokens[i].Text;
            if (i == 0 && tokenText == "<<") continue;
            if (tokenText.Length > 0 && tokenText[0] == '/') {
                string key = DecodeName(tokenText.Substring(1));
                if (i + 1 < tokens.Count && tokens[i + 1].Text != ">>") {
                    var (obj, consumed) = ParseObject(tokens, i + 1, effectiveLimits, 0);
                    d.HasIncompleteSyntax |= d.Items.ContainsKey(key);
                    d.Items[key] = obj;
                    d.HasIncompleteSyntax |= obj.HasIncompleteSyntax;
                    i += consumed + 1;
                } else d.HasIncompleteSyntax = true;
            } else if (tokenText == ">>" && i == tokens.Count - 1) {
                break;
            } else d.HasIncompleteSyntax = true;
        }
        return d;
    }

    private static (PdfObject Obj, int Consumed) ParseObject(in PooledTokenBuffer tokens, int i, PdfReadLimits limits, int depth) {
        if (i < 0 || i >= tokens.Count) return (new PdfName(""), 0);
        if (depth > limits.MaxObjectNestingDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, limits.MaxObjectNestingDepth, depth);
        }

        PdfToken token = tokens[i];
        string tok = token.Text;
        if (tok == "<<") {
            var dict = new PdfDictionary();
            int j = i + 1;
            while (j < tokens.Count && tokens[j].Text != ">>") {
                string keyToken = tokens[j].Text;
                if (keyToken.Length > 0 && keyToken[0] == '/') {
                    string key = DecodeName(keyToken.Substring(1));
                    if (j + 1 < tokens.Count && tokens[j + 1].Text != ">>") {
                        var (obj, consumed) = ParseObject(tokens, j + 1, limits, depth + 1);
                        dict.HasIncompleteSyntax |= dict.Items.ContainsKey(key);
                        dict.Items[key] = obj;
                        dict.HasIncompleteSyntax |= obj.HasIncompleteSyntax;
                        j += consumed + 2;
                        continue;
                    }
                }
                dict.HasIncompleteSyntax = true;
                j++;
            }
            dict.HasIncompleteSyntax |= j >= tokens.Count;
            return (dict, j - i);
        }
        if (tok == "[") {
            var arr = new PdfArray(); int j = i + 1;
            while (j < tokens.Count && tokens[j].Text != "]") {
                var (inner, used) = ParseObject(tokens, j, limits, depth + 1);
                arr.Items.Add(inner);
                arr.HasIncompleteSyntax |= inner.HasIncompleteSyntax;
                j += used + 1;
            }
            arr.HasIncompleteSyntax |= j >= tokens.Count;
            return (arr, j - i);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(DecodeName(tok.Substring(1))), 0);
        if (token.IsString && tok.Length > 0 && tok[0] == '(') {
            bool isTerminated = token.IsTerminated;
            string inner = isTerminated
                ? tok.Substring(1, tok.Length - 2)
                : tok.Substring(1);
            var value = CreateParsedString(
                PdfStringParser.ParseLiteralToBytes(inner),
                token.EncodedLength);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (token.IsString && tok.Length > 0 && tok[0] == '<' && (tok.Length == 1 || tok[1] != '<')) {
            bool isTerminated = token.IsTerminated;
            string inner = isTerminated
                ? tok.Substring(1, tok.Length - 2)
                : tok.Substring(1);
            var value = CreateParsedString(
                PdfTextString.DecodeHexBytes(inner),
                token.EncodedLength);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (string.Equals(tok, "true", StringComparison.Ordinal)) return (new PdfBoolean(true), 0);
        if (string.Equals(tok, "false", StringComparison.Ordinal)) return (new PdfBoolean(false), 0);
        if (string.Equals(tok, "null", StringComparison.Ordinal)) return (PdfNull.Instance, 0);
        if (tok.Length > 0 && (char.IsDigit(tok[0]) || tok[0] == '-' || tok[0] == '+' || tok[0] == '.')) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2].Text == "R" && int.TryParse(tokens[i].Text, out int obj) && int.TryParse(tokens[i + 1].Text, out int gen)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (double.TryParse(tok, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(tok) { HasIncompleteSyntax = true }, 0);
    }

    private static PooledTokenBuffer Tokenize(
        string s,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) =>
        Tokenize(s, 0, s.Length, limits, trackEncodedStringSourceSpans);

    private static PooledTokenBuffer Tokenize(
        string s,
        int start,
        int length,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) {
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, length);
        }
        if (start < 0 || length < 0 || start > s.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        int end = start + length;

        // Object length is a poor token-count estimate when a dictionary holds
        // a long literal or hex string. Grow for genuinely dense objects rather
        // than reserving thousands of unused token slots up front.
        int estimatedTokens = Math.Min(512, length / 4 + 8);
        var tokens = new PooledTokenBuffer(Math.Min(estimatedTokens, effectiveLimits.MaxTokensPerObject));
        try {
            TokenizeInto(s, start, end, effectiveLimits, trackEncodedStringSourceSpans, ref tokens);
            return tokens;
        } catch {
            tokens.Dispose();
            throw;
        }
    }

    private static void TokenizeInto(
        string s,
        int start,
        int end,
        PdfReadLimits effectiveLimits,
        bool trackEncodedStringSourceSpans,
        ref PooledTokenBuffer tokens) {
        int i = start;
        while (i < end) {
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }

            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '%') {
                i++;
                while (i < end && s[i] != '\n' && s[i] != '\r') i++;
                continue;
            }
            if (c == '<' && i + 1 < end && s[i + 1] == '<') { tokens.Add(new PdfToken("<<")); i += 2; continue; }
            if (c == '>' && i + 1 < end && s[i + 1] == '>') { tokens.Add(new PdfToken(">>")); i += 2; continue; }
            if (c == '[' || c == ']') { tokens.Add(new PdfToken(c.ToString())); i++; continue; }
            if (c == '<') {
                int tokenStart = i++;
                while (i < end && s[i] != '>') i++;
                if (i < end && s[i] == '>') i++;
                bool isTerminated = i > tokenStart && s[i - 1] == '>';
                string text = s.Substring(tokenStart, i - tokenStart);
                tokens.Add(new PdfToken(
                    text,
                    isString: true,
                    isTerminated,
                    isTerminated && trackEncodedStringSourceSpans ? text.Length : null));
                continue;
            }
            if (c == '(') {
                int tokenStart = i; i++;
                int depth = 1; bool esc = false;
                while (i < end && depth > 0) {
                    char ch = s[i++];
                    if (esc) { esc = false; } else if (ch == '\\') { esc = true; } else if (ch == '(') {
                        depth++;
                        if (depth > effectiveLimits.MaxObjectNestingDepth) {
                            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, effectiveLimits.MaxObjectNestingDepth, depth);
                        }
                    } else if (ch == ')') { depth--; }
                }
                // The scan only finds the token boundary; its raw bytes already
                // contain the exact nested parentheses and escape spelling.
                string text = s.Substring(tokenStart, i - tokenStart);
                bool isTerminated = depth == 0;
                tokens.Add(new PdfToken(
                    text,
                    isString: true,
                    isTerminated,
                    isTerminated && trackEncodedStringSourceSpans ? text.Length : null));
                continue;
            }
            // name, number, keyword
            int j = i;
            while (j < end && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
            string tok = s.Substring(i, j - i);
            if (tok.Length == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < end && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tok = s.Substring(i, j - i);
            }
            if (tok.Length == 0) {
                // A malformed or unexpected standalone delimiter must still consume input.
                // Otherwise ')' and '>' repeatedly produce empty tokens until the token
                // budget is exhausted, turning one bad byte into excessive CPU and memory use.
                tok = s[i].ToString();
                j = i + 1;
            }
            tokens.Add(new PdfToken(tok));
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }
            i = j;
        }

        if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
        }
    }

    /// <summary>Keeps small object token lists local and rents large backing arrays for dense PDF objects.</summary>
    private struct PooledTokenBuffer : IDisposable {
        private const int SmallTokenLimit = 1024;
        private List<PdfToken>? _small;
#if NET8_0_OR_GREATER
        private PdfToken[]? _rented;
#endif
        private int _count;

        internal PooledTokenBuffer(int initialCapacity) {
            _small = new List<PdfToken>(initialCapacity);
        }

        internal readonly int Count => _count;

        internal readonly PdfToken this[int index] {
            get {
                if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index));
#if NET8_0_OR_GREATER
                return _rented is null ? _small![index] : _rented[index];
#else
                return _small![index];
#endif
            }
        }

        internal void Add(PdfToken token) {
#if NET8_0_OR_GREATER
            if (_rented is null && _count < SmallTokenLimit) {
                _small!.Add(token);
            } else {
                if (_rented is null) {
                    _rented = ArrayPool<PdfToken>.Shared.Rent(SmallTokenLimit * 2);
                    _small!.CopyTo(_rented);
                    _small = null;
                } else if (_count == _rented.Length) {
                    PdfToken[] expanded = ArrayPool<PdfToken>.Shared.Rent(checked(_count * 2));
                    Array.Copy(_rented, expanded, _count);
                    ArrayPool<PdfToken>.Shared.Return(_rented, clearArray: true);
                    _rented = expanded;
                }

                _rented[_count] = token;
            }
#else
            _small!.Add(token);
#endif

            _count++;
        }

        public void Dispose() {
#if NET8_0_OR_GREATER
            if (_rented is not null) {
                ArrayPool<PdfToken>.Shared.Return(_rented, clearArray: true);
                _rented = null;
            }
#endif
        }
    }

    private readonly struct PdfToken {
        internal PdfToken(
            string text,
            bool isString = false,
            bool isTerminated = false,
            int? encodedLength = null) {
            Text = text ?? string.Empty;
            IsString = isString;
            IsTerminated = isTerminated;
            EncodedLength = encodedLength;
        }

        internal string Text { get; }
        internal bool IsString { get; }
        internal bool IsTerminated { get; }
        internal int? EncodedLength { get; }
    }

    private static bool TryGetResolvedLength(PdfDictionary dict, Dictionary<int, PdfIndirectObject> map, out int length) {
        length = -1;

        if (dict.Get<PdfNumber>("Length") is PdfNumber lenNum) {
            int resolved = (int)Math.Max(0, Math.Min(int.MaxValue, lenNum.Value));
            length = resolved;
            return true;
        }

        if (dict.Get<PdfReference>("Length") is PdfReference lenRef &&
            map.TryGetValue(lenRef.ObjectNumber, out var indirectLength) &&
            indirectLength.Value is PdfNumber referencedLength) {
            int resolved = (int)Math.Max(0, Math.Min(int.MaxValue, referencedLength.Value));
            length = resolved;
            return true;
        }

        return false;
    }

    internal static string DecodeName(string raw) {
        if (string.IsNullOrEmpty(raw) || raw.IndexOf('#') < 0) {
            return raw;
        }

        var sb = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++) {
            char ch = raw[i];
            if (ch == '#' && i + 2 < raw.Length && TryHexNibble(raw[i + 1], out int hi) && TryHexNibble(raw[i + 2], out int lo)) {
                sb.Append(PdfEncoding.Latin1GetString(new[] { (byte)((hi << 4) | lo) }));
                i += 2;
                continue;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }

    private static PdfStringObj CreateParsedString(byte[] bytes, int? encodedTokenLength) {
        string value = PdfTextString.Decode(bytes);
        return PdfStringObj.FromParsedBytes(
            bytes,
            value,
            useTextStringEncoding: !PdfWinAnsiEncoding.CanEncode(value, out _),
            encodedTokenLength);
    }

    private static bool TryHexNibble(char c, out int value) {
        if (c >= '0' && c <= '9') {
            value = c - '0';
            return true;
        }
        if (c >= 'a' && c <= 'f') {
            value = 10 + (c - 'a');
            return true;
        }
        if (c >= 'A' && c <= 'F') {
            value = 10 + (c - 'A');
            return true;
        }

        value = 0;
        return false;
    }
}
