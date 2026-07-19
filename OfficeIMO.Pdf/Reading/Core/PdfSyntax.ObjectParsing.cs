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
                    int dictionaryCharacters = dictEnd - (dictStart + 2);
                    if (dictionaryCharacters > effectiveLimits.MaxObjectCharacters) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, dictionaryCharacters);
                    }

                    string dictText = SafeSlice(body, dictStart + 2, dictionaryCharacters, effectiveLimits.MaxObjectCharacters);
                    try { return ParseDictionary(dictText, effectiveLimits, trackEncodedStringSourceSpans); }
                    catch (PdfReadLimitException) { throw; }
                    catch { return null; }
                }
            }
            return null;
        }
        if (s.Length > 0 && s[0] == '[') {
            var toks = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, _) = ParseObject(toks, 0, effectiveLimits, 0);
            return obj;
        }
        if (s.Length > 0 && s[0] == '(') {
            var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, _) = ParseObject(stringTokens, 0, effectiveLimits, 0);
            return obj;
        }
        if (s.Length > 0 && s[0] == '<' && (s.Length == 1 || s[1] != '<')) {
            var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
            var (obj, _) = ParseObject(stringTokens, 0, effectiveLimits, 0);
            return obj;
        }
        // number or name fallbacks
        var tokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans);
        if (tokens.Count > 0) {
            var (obj0, _) = ParseObject(tokens, 0, effectiveLimits, 0);
            return obj0;
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
        bool trackEncodedStringSourceSpans = true) {
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (dict.Length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, dict.Length);
        }

        var d = new PdfDictionary();
        var tokens = Tokenize(dict, effectiveLimits, trackEncodedStringSourceSpans);
        for (int i = 0; i < tokens.Count; i++) {
            string tokenText = tokens[i].Text;
            if (tokenText.Length > 0 && tokenText[0] == '/') {
                string key = DecodeName(tokenText.Substring(1));
                if (i + 1 < tokens.Count) {
                    var (obj, consumed) = ParseObject(tokens, i + 1, effectiveLimits, 0);
                    d.Items[key] = obj;
                    i += consumed + 1;
                }
            }
        }
        return d;
    }

    private static (PdfObject Obj, int Consumed) ParseObject(List<PdfToken> tokens, int i, PdfReadLimits limits, int depth) {
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
                    if (j + 1 < tokens.Count) {
                        var (obj, consumed) = ParseObject(tokens, j + 1, limits, depth + 1);
                        dict.Items[key] = obj;
                        j += consumed + 2;
                        continue;
                    }
                }
                j++;
            }
            return (dict, j - i);
        }
        if (tok == "[") {
            var arr = new PdfArray(); int j = i + 1;
            while (j < tokens.Count && tokens[j].Text != "]") {
                var (inner, used) = ParseObject(tokens, j, limits, depth + 1);
                arr.Items.Add(inner);
                j += used + 1;
            }
            return (arr, j - i);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(DecodeName(tok.Substring(1))), 0);
        if (token.IsString && tok.Length > 0 && tok[0] == '(') {
            bool isTerminated = token.IsTerminated;
            string inner = isTerminated
                ? tok.Substring(1, tok.Length - 2)
                : tok.Substring(1);
            return (CreateParsedString(
                PdfStringParser.ParseLiteralToBytes(inner),
                token.EncodedLength), 0);
        }
        if (token.IsString && tok.Length > 0 && tok[0] == '<' && (tok.Length == 1 || tok[1] != '<')) {
            bool isTerminated = token.IsTerminated;
            string inner = isTerminated
                ? tok.Substring(1, tok.Length - 2)
                : tok.Substring(1);
            return (CreateParsedString(
                PdfTextString.DecodeHexBytes(inner),
                token.EncodedLength), 0);
        }
        if (string.Equals(tok, "true", StringComparison.Ordinal)) return (new PdfBoolean(true), 0);
        if (string.Equals(tok, "false", StringComparison.Ordinal)) return (new PdfBoolean(false), 0);
        if (string.Equals(tok, "null", StringComparison.Ordinal)) return (PdfNull.Instance, 0);
        if (tok.Length > 0 && (char.IsDigit(tok[0]) || tok[0] == '-' || tok[0] == '+')) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2].Text == "R" && int.TryParse(tokens[i].Text, out int obj) && int.TryParse(tokens[i + 1].Text, out int gen)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (double.TryParse(tok, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(tok), 0);
    }

    private static List<PdfToken> Tokenize(
        string s,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true) {
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (s.Length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, s.Length);
        }

        var tokens = new List<PdfToken>(Math.Min(16384, s.Length / 2 + 8));
        int i = 0;
        while (i < s.Length) {
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }

            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '%') {
                i++;
                while (i < s.Length && s[i] != '\n' && s[i] != '\r') i++;
                continue;
            }
            if (c == '<' && i + 1 < s.Length && s[i + 1] == '<') { tokens.Add(new PdfToken("<<")); i += 2; continue; }
            if (c == '>' && i + 1 < s.Length && s[i + 1] == '>') { tokens.Add(new PdfToken(">>")); i += 2; continue; }
            if (c == '[' || c == ']') { tokens.Add(new PdfToken(c.ToString())); i++; continue; }
            if (c == '<') {
                int start = i++;
                while (i < s.Length && s[i] != '>') i++;
                if (i < s.Length && s[i] == '>') i++;
                bool isTerminated = i > start && s[i - 1] == '>';
                string text = s.Substring(start, i - start);
                tokens.Add(new PdfToken(
                    text,
                    isString: true,
                    isTerminated,
                    isTerminated && trackEncodedStringSourceSpans ? text.Length : null));
                continue;
            }
            if (c == '(') {
                int start = i; i++;
                int depth = 1; bool esc = false;
                var sb = new StringBuilder();
                while (i < s.Length && depth > 0) {
                    char ch = s[i++];
                    if (esc) { sb.Append(ch); esc = false; } else if (ch == '\\') { sb.Append(ch); esc = true; }
                    else if (ch == '(') {
                        depth++;
                        if (depth > effectiveLimits.MaxObjectNestingDepth) {
                            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, effectiveLimits.MaxObjectNestingDepth, depth);
                        }

                        sb.Append(ch);
                    } else if (ch == ')') { depth--; if (depth > 0) sb.Append(ch); } else sb.Append(ch);
                }
                string text = depth == 0
                    ? "(" + sb.ToString() + ")"
                    : "(" + sb.ToString();
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
            while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
            string tok = s.Substring(i, j - i);
            if (tok.Length == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tok = s.Substring(i, j - i);
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

        return tokens;
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
        return new PdfStringObj(
            bytes,
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
