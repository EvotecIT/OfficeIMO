#if NET8_0_OR_GREATER
using System.Buffers;
using System.Runtime.InteropServices;
#endif
namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    // Numeric values beyond this length cannot add useful PDF numeric precision and
    // must not enter tokenless framework parsing after a caller raises object limits.
    private const int MaxNumericTokenCharacters = 4096;
    private static readonly Encoding StrictUtf8NameEncoding = new UTF8Encoding(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);

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
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
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
                int dictEnd = FindDictEnd(body, dictStart, body.Length, cancellationToken);
                if (dictEnd > dictStart) {
                    if (SkipWhitespaceAndComments(body, dictEnd, body.Length, cancellationToken) < body.Length) return null;
                    int dictionaryCharacters = dictEnd - (dictStart + 2);
                    if (dictionaryCharacters > effectiveLimits.MaxObjectCharacters) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, dictionaryCharacters);
                    }

                    try { return ParseDictionary(body, dictStart + 2, dictionaryCharacters, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken); }
                    catch (Exception ex) when (ex is not PdfReadLimitException and not OperationCanceledException and not OutOfMemoryException) { return null; }
                }
            }
            return null;
        }
        if (s.Length > 0 && s[0] == '[') {
            using var toks = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
            var (obj, consumed) = ParseObject(toks, 0, effectiveLimits, 0, cancellationToken);
            return consumed + 1 == toks.Count ? obj : null;
        }
        if (s.Length > 0 && s[0] == '(') {
            using var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
            var (obj, consumed) = ParseObject(stringTokens, 0, effectiveLimits, 0, cancellationToken);
            return consumed + 1 == stringTokens.Count ? obj : null;
        }
        if (s.Length > 0 && s[0] == '<' && (s.Length == 1 || s[1] != '<')) {
            using var stringTokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
            var (obj, consumed) = ParseObject(stringTokens, 0, effectiveLimits, 0, cancellationToken);
            return consumed + 1 == stringTokens.Count ? obj : null;
        }
        // number or name fallbacks
        using var tokens = Tokenize(s, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
        if (tokens.Count > 0) {
            var (obj0, consumed) = ParseObject(tokens, 0, effectiveLimits, 0, cancellationToken);
            return consumed + 1 == tokens.Count ? obj0 : null;
        }
        return null;
    }

    /// <summary>
    /// Parses a top-level object held in a window of decoded bytes, such as one object of an object
    /// stream, with the same results as <see cref="ParseTopLevelObject(string, PdfReadLimits?, bool, System.Threading.CancellationToken)"/>
    /// over the Latin-1 string of that window, without building the string.
    /// </summary>
    private static PdfObject? ParseTopLevelObject(
        byte[] data,
        int start,
        int length,
        PdfReadLimits? limits,
        bool trackEncodedStringSourceSpans,
        System.Threading.CancellationToken cancellationToken) =>
        ParseTopLevelObject(new PdfByteContentText(data), start, length, limits, trackEncodedStringSourceSpans, cancellationToken);

    private static PdfObject? ParseTopLevelObject<TText>(
        TText text,
        int start,
        int length,
        PdfReadLimits? limits,
        bool trackEncodedStringSourceSpans,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        int end = start + length;
        int first = start;
        while (first < end && char.IsWhiteSpace(text[first])) first++;
        if (first == end) return null;
        if (length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, length);
        }

        int trimmedLength = end - first;
        if (IsWindow(text, first, end, "true")) return new PdfBoolean(true);
        if (IsWindow(text, first, end, "false")) return new PdfBoolean(false);
        if (IsWindow(text, first, end, "null")) return PdfNull.Instance;
        if (trimmedLength >= 2 && text[first] == '<' && text[first + 1] == '<') {
            // Leading whitespace holds no '<', so the first "<<" of the window starts the trimmed text.
            int dictStart = first;
            int dictEnd = FindDictEnd(text, dictStart, end, cancellationToken);
            if (dictEnd > dictStart) {
                if (SkipWhitespaceAndComments(text, dictEnd, end, cancellationToken) < end) return null;
                int dictionaryCharacters = dictEnd - (dictStart + 2);
                if (dictionaryCharacters > effectiveLimits.MaxObjectCharacters) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, dictionaryCharacters);
                }

                try { return ParseDictionary(text, dictStart + 2, dictionaryCharacters, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken); }
                catch (Exception ex) when (ex is not PdfReadLimitException and not OperationCanceledException and not OutOfMemoryException) { return null; }
            }
            return null;
        }

        using var tokens = Tokenize(text, first, trimmedLength, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
        if (tokens.Count == 0) return null;
        try {
            var (obj, consumed) = ParseObject(tokens, 0, effectiveLimits, 0, cancellationToken);
            return consumed + 1 == tokens.Count ? obj : null;
        } catch (FormatException) {
            // A malformed hex string makes the object unreadable, as it does for dictionaries above.
            return null;
        }
    }

    private static bool IsWindow<TText>(TText text, int start, int end, string value) where TText : struct, IPdfContentText {
        if (end - start != value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (text[start + index] != value[index]) return false;
        }
        return true;
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
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) =>
        ParseDictionary(dict, 0, dict.Length, limits, trackEncodedStringSourceSpans, cancellationToken);

    private static PdfDictionary ParseDictionary(
        string source,
        int start,
        int length,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) =>
        ParseDictionary(new PdfStringContentText(source), start, length, limits, trackEncodedStringSourceSpans, cancellationToken);

    private static PdfDictionary ParseDictionary<TText>(
        TText source,
        int start,
        int length,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        if (length > effectiveLimits.MaxObjectCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectCharacters, effectiveLimits.MaxObjectCharacters, length);
        }

        using var tokens = Tokenize(source, start, length, effectiveLimits, trackEncodedStringSourceSpans, cancellationToken);
        // Object bodies usually arrive without their opening delimiter. Size them like a nested dictionary.
        var d = new PdfDictionary(EstimateDictionaryKeyCount(tokens,
            tokens.Count > 0 && tokens[0].Equals(tokens.Source, "<<") ? 0 : -1, cancellationToken));
        for (int i = 0; i < tokens.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfToken token = tokens[i];
            if (i == 0 && token.Equals(tokens.Source, "<<")) continue;
            if (token.Length > 0 && token.GetCharacter(tokens.Source, 0) == '/') {
                string key = DecodeNameToken(token, tokens.Source, cancellationToken);
                if (i + 1 < tokens.Count && !tokens[i + 1].Equals(tokens.Source, ">>")) {
                    var (obj, consumed) = ParseObject(tokens, i + 1, effectiveLimits, 0, cancellationToken);
                    SetDictionaryItem(d, key, obj);
                    i += consumed + 1;
                } else d.HasIncompleteSyntax = true;
            } else if (token.Equals(tokens.Source, ">>") && i == tokens.Count - 1) {
                break;
            } else d.HasIncompleteSyntax = true;
        }
        return d;
    }

    private static (PdfObject Obj, int Consumed) ParseObject<TText>(in PooledTokenBuffer<TText> tokens, int i, PdfReadLimits limits, int depth,
        System.Threading.CancellationToken cancellationToken = default) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
        if (i < 0 || i >= tokens.Count) return (new PdfName(""), 0);
        if (depth > limits.MaxObjectNestingDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, limits.MaxObjectNestingDepth, depth);
        }

        PdfToken token = tokens[i];
        TText source = tokens.Source;
        if (token.Equals(source, "<<")) {
            var dict = new PdfDictionary(EstimateDictionaryKeyCount(tokens, i, cancellationToken));
            int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, ">>")) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfToken keyToken = tokens[j];
                if (keyToken.Length > 0 && keyToken.GetCharacter(source, 0) == '/') {
                    string key = DecodeNameToken(keyToken, source, cancellationToken);
                    if (j + 1 < tokens.Count && !tokens[j + 1].Equals(source, ">>")) {
                        var (obj, consumed) = ParseObject(tokens, j + 1, limits, depth + 1, cancellationToken);
                        SetDictionaryItem(dict, key, obj);
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
        if (token.Equals(source, "[")) {
            var arr = new PdfArray(EstimateArrayItemCount(tokens, i, cancellationToken)); int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, "]")) {
                cancellationToken.ThrowIfCancellationRequested();
                var (inner, used) = ParseObject(tokens, j, limits, depth + 1, cancellationToken);
                arr.Items.Add(inner);
                arr.HasIncompleteSyntax |= inner.HasIncompleteSyntax;
                j += used + 1;
            }
            arr.HasIncompleteSyntax |= j >= tokens.Count;
            return (arr, j - i);
        }
        if (token.Length > 0 && token.GetCharacter(source, 0) == '/') return (new PdfName(DecodeNameToken(token, source, cancellationToken)), 0);
        if (token.IsString && token.Length > 0 && token.GetCharacter(source, 0) == '(') {
            bool isTerminated = token.IsTerminated;
            int innerLength = token.Length - (isTerminated ? 2 : 1);
            var value = CreateParsedString(
                PdfStringParser.ParseLiteralToBytes(source, token.SourceStart + 1, innerLength, cancellationToken),
                token.EncodedLength,
                cancellationToken);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (token.IsString && token.Length > 0 && token.GetCharacter(source, 0) == '<' &&
            (token.Length == 1 || token.GetCharacter(source, 1) != '<')) {
            bool isTerminated = token.IsTerminated;
            int innerLength = token.Length - (isTerminated ? 2 : 1);
            var value = CreateParsedString(
                PdfTextString.DecodeHexBytes(source, token.SourceStart + 1, innerLength, cancellationToken),
                token.EncodedLength,
                cancellationToken);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (token.Equals(source, "true")) return (new PdfBoolean(true), 0);
        if (token.Equals(source, "false")) return (new PdfBoolean(false), 0);
        if (token.Equals(source, "null")) return (PdfNull.Instance, 0);
        if (token.Length > 0 && IsPdfNumberStart(token.GetCharacter(source, 0))) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2].Equals(source, "R") &&
                token.TryParseInt32(source, out int obj, cancellationToken) && tokens[i + 1].TryParseInt32(source, out int gen, cancellationToken)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (token.TryParseDouble(source, out double val, cancellationToken)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(token.GetText(source, cancellationToken)) { HasIncompleteSyntax = true }, 0);
    }

    // Counts the keys of a dictionary whose entries are found within a short scan, so it is
    // allocated at its final size. Long or malformed dictionaries return 0 and grow as parsed.
    private static int EstimateDictionaryKeyCount<TText>(in PooledTokenBuffer<TText> tokens, int dictionaryStart,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        const int MaximumScannedTokens = 64;
        TText source = tokens.Source;
        int end = Math.Min(tokens.Count, dictionaryStart + 1 + MaximumScannedTokens);
        int count = 0;
        int depth = 0;
        bool expectKey = true;
        for (int index = dictionaryStart + 1; index < end; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfToken token = tokens[index];
            if (token.Equals(source, "<<") || token.Equals(source, "[")) {
                depth++;
                continue;
            }
            if (token.Equals(source, ">>") || token.Equals(source, "]")) {
                if (depth == 0) return token.Equals(source, ">>") ? count : 0;
                if (--depth == 0) expectKey = true;
                continue;
            }
            if (depth > 0) continue;
            if (expectKey) {
                if (token.Length == 0 || token.GetCharacter(source, 0) != '/') return 0;
                count++;
                expectKey = false;
            } else {
                // The count only sizes the dictionary, so a reference is recognized by its shape without parsing it.
                if (index + 2 < end &&
                    tokens[index + 2].Equals(source, "R") &&
                    StartsWithDigit(tokens[index], source) &&
                    StartsWithDigit(tokens[index + 1], source)) {
                    index += 2;
                }
                expectKey = true;
            }
        }
        return 0;
    }

    private static int EstimateArrayItemCount<TText>(in PooledTokenBuffer<TText> tokens, int arrayStart,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        int count = 0;
        int index = arrayStart + 1;
        TText source = tokens.Source;
        while (index < tokens.Count) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfToken token = tokens[index];
            if (token.Equals(source, "]")) break;
            count++;
            if (token.Equals(source, "[") || token.Equals(source, "<<")) {
                // Do not rescan nested subtrees. Their own ParseObject call can size
                // a flat child array, while the parent grows from observed entries.
                return 0;
            }

            if (index + 2 < tokens.Count &&
                tokens[index + 2].Equals(source, "R") &&
                StartsWithDigit(tokens[index], source) &&
                StartsWithDigit(tokens[index + 1], source)) {
                index += 3;
                continue;
            }

            index++;
        }

        return count;
    }

    private static bool StartsWithDigit<TText>(PdfToken token, TText source) where TText : struct, IPdfContentText =>
        !token.IsString && token.Length > 0 && (uint)(token.GetCharacter(source, 0) - '0') <= 9;

    private static PooledTokenBuffer<PdfStringContentText> Tokenize(
        string s,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) =>
        Tokenize(new PdfStringContentText(s), 0, s.Length, limits, trackEncodedStringSourceSpans, cancellationToken);

    private static PooledTokenBuffer<PdfStringContentText> Tokenize(
        string s,
        int start,
        int length,
        PdfReadLimits? limits = null,
        bool trackEncodedStringSourceSpans = true,
        System.Threading.CancellationToken cancellationToken = default) =>
        Tokenize(new PdfStringContentText(s), start, length, limits, trackEncodedStringSourceSpans, cancellationToken);

    private static PooledTokenBuffer<TText> Tokenize<TText>(
        TText s,
        int start,
        int length,
        PdfReadLimits? limits,
        bool trackEncodedStringSourceSpans,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
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
        var tokens = new PooledTokenBuffer<TText>(s, Math.Min(estimatedTokens, effectiveLimits.MaxTokensPerObject));
        try {
            TokenizeInto(s, start, end, effectiveLimits, trackEncodedStringSourceSpans, ref tokens, cancellationToken);
            return tokens;
        } catch {
            tokens.Dispose();
            throw;
        }
    }

    private static void TokenizeInto<TText>(
        TText s,
        int start,
        int end,
        PdfReadLimits effectiveLimits,
        bool trackEncodedStringSourceSpans,
        ref PooledTokenBuffer<TText> tokens,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        int i = start;
        while (i < end) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }

            char c = s[i];
            if (c < 256 ? WhiteSpaceCharacters[c] : char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '%') {
                i++;
                while (i < end && s[i] != '\n' && s[i] != '\r') {
                    if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    i++;
                }
                continue;
            }
            if (c == '<' && i + 1 < end && s[i + 1] == '<') { tokens.Add(new PdfToken("<<")); i += 2; continue; }
            if (c == '>' && i + 1 < end && s[i + 1] == '>') { tokens.Add(new PdfToken(">>")); i += 2; continue; }
            if (c == '[' || c == ']') { tokens.Add(new PdfToken(c == '[' ? "[" : "]")); i++; continue; }
            if (c == '<') {
                int tokenStart = i++;
                while (i < end && s[i] != '>') {
                    if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    i++;
                }
                if (i < end && s[i] == '>') i++;
                bool isTerminated = i > tokenStart && s[i - 1] == '>';
                tokens.Add(new PdfToken(
                    tokenStart,
                    i - tokenStart,
                    isString: true,
                    isTerminated,
                    isTerminated && trackEncodedStringSourceSpans ? i - tokenStart : null));
                continue;
            }
            if (c == '(') {
                int tokenStart = i; i++;
                int depth = 1; bool esc = false;
                while (i < end && depth > 0) {
                    if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
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
                bool isTerminated = depth == 0;
                tokens.Add(new PdfToken(
                    tokenStart,
                    i - tokenStart,
                    isString: true,
                    isTerminated,
                    isTerminated && trackEncodedStringSourceSpans ? i - tokenStart : null));
                continue;
            }
            // name, number, keyword
            int j = i;
            while (j < end && !EndsBareToken(s[j])) {
                if ((j & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                j++;
            }
            int tokenLength = j - i;
            if (tokenLength == 0 && c == '/') { // name starting here
                j = i + 1;
                while (j < end && !EndsBareToken(s[j])) {
                    if ((j & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    j++;
                }
                tokenLength = j - i;
            }
            if (tokenLength == 0) {
                // A malformed or unexpected standalone delimiter must still consume input.
                // Otherwise ')' and '>' repeatedly produce empty tokens until the token
                // budget is exhausted, turning one bad byte into excessive CPU and memory use.
                tokenLength = 1;
                j = i + 1;
            }
            tokens.Add(MaterializeToken(s, i, tokenLength, cancellationToken));
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }
            i = j;
        }

        if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
        }
    }

    // Whitespace (as char.IsWhiteSpace defines it) and the delimiters that end a name, number or keyword token.
    private static readonly bool[] WhiteSpaceCharacters = CreateWhiteSpaceCharacters();
    private static readonly bool[] BareTokenTerminators = CreateBareTokenTerminators();

    private static bool[] CreateWhiteSpaceCharacters() {
        var table = new bool[256];
        for (int value = 0; value < table.Length; value++) table[value] = char.IsWhiteSpace((char)value);
        return table;
    }

    private static bool[] CreateBareTokenTerminators() {
        bool[] table = CreateWhiteSpaceCharacters();
        foreach (char delimiter in "%/[]<>()") table[delimiter] = true;
        return table;
    }

    private static bool EndsBareToken(char value) =>
        value < 256 ? BareTokenTerminators[value] : char.IsWhiteSpace(value);

    /// <summary>Reuses token arrays on modern targets while retaining the compatible list path on legacy targets.</summary>
    private struct PooledTokenBuffer<TText> : IDisposable where TText : struct, IPdfContentText {
        private PdfToken[]? _rented;
        private int _count;
        private readonly TText _source;

        internal PooledTokenBuffer(TText source, int initialCapacity) {
            _source = source;
            _rented = RentTokens(Math.Max(initialCapacity, 16));
        }

        internal readonly int Count => _count;
        internal readonly TText Source => _source;

        internal readonly PdfToken this[int index] {
            get {
                if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index));
                return _rented![index];
            }
        }

        internal void Add(PdfToken token) {
            if (_count == _rented!.Length) {
                PdfToken[] expanded = RentTokens(checked(_count * 2));
                Array.Copy(_rented, expanded, _count);
                ReturnTokens(_rented, _count);
                _rented = expanded;
            }

            _rented[_count] = token;

            _count++;
        }

        public void Dispose() {
            if (_rented is not null) {
                ReturnTokens(_rented, _count);
                _rented = null;
            }
        }
    }

#if NET8_0_OR_GREATER
    // Token arrays above this size are large-object-heap allocations. The shared pool drops its
    // arrays on every gen2 collection, so a document with one dense object (a structure parent
    // tree, say) allocated a fresh large array on each load. Each thread keeps its largest
    // array instead; a nested tokenization while it is lent out uses the shared pool.
    private const int LargeTokenArrayLength = 2048;
    // Bounds what an idle thread retains after an unusually dense object.
    private const int MaximumCachedTokenArrayLength = 1 << 17;
    [ThreadStatic] private static PdfToken[]? t_largeTokens;

    private static PdfToken[] RentTokens(int minimumLength) {
        if (minimumLength > LargeTokenArrayLength) {
            PdfToken[]? cached = t_largeTokens;
            // A much larger array would make every return clear far more slots than this object used.
            if (cached is not null && cached.Length >= minimumLength && cached.Length / 4 <= minimumLength) {
                t_largeTokens = null;
                return cached;
            }
        }
        return ArrayPool<PdfToken>.Shared.Rent(minimumLength);
    }

    private static void ReturnTokens(PdfToken[] tokens, int _) => ReturnTokens(tokens);

    private static void ReturnTokens(PdfToken[] tokens) {
        if (tokens.Length > LargeTokenArrayLength && tokens.Length <= MaximumCachedTokenArrayLength) {
            PdfToken[]? cached = t_largeTokens;
            if (cached is null || cached.Length < tokens.Length) {
                // Tokens can reference source strings; clear them before keeping the array.
                Array.Clear(tokens);
                t_largeTokens = tokens;
                if (cached is not null) ArrayPool<PdfToken>.Shared.Return(cached, clearArray: true);
                return;
            }
        }
        ArrayPool<PdfToken>.Shared.Return(tokens, clearArray: true);
    }
#else
    // Older targets have no shared array pool, and a new token list per object was a measurable share of loading.
    // Each thread keeps its largest token array; a nested tokenization while it is lent out allocates its own.
    private const int MaximumCachedTokenArrayLength = 1 << 17;
    [ThreadStatic] private static PdfToken[]? t_tokens;

    private static PdfToken[] RentTokens(int minimumLength) {
        PdfToken[]? cached = t_tokens;
        if (cached is not null && cached.Length >= minimumLength) {
            t_tokens = null;
            return cached;
        }
        return new PdfToken[minimumLength];
    }

    private static void ReturnTokens(PdfToken[] tokens, int used) {
        if (tokens.Length > MaximumCachedTokenArrayLength) return;
        PdfToken[]? cached = t_tokens;
        if (cached is not null && cached.Length >= tokens.Length) return;
        // Tokens can reference source strings; clear the used slots before keeping the array.
        Array.Clear(tokens, 0, Math.Min(used, tokens.Length));
        t_tokens = tokens;
    }
#endif

    private readonly struct PdfToken {
        private readonly string? _text;
        private readonly int _sourceStart;
        private readonly int _length;

        internal PdfToken(
            string text,
            bool isString = false,
            bool isTerminated = false,
            int? encodedLength = null) {
            _text = text ?? string.Empty;
            _sourceStart = -1;
            _length = _text.Length;
            IsString = isString;
            IsTerminated = isTerminated;
            EncodedLength = encodedLength;
        }

        internal PdfToken(
            int sourceStart,
            int length,
            bool isString = false,
            bool isTerminated = false,
            int? encodedLength = null) {
            _text = null;
            _sourceStart = sourceStart;
            _length = length;
            IsString = isString;
            IsTerminated = isTerminated;
            EncodedLength = encodedLength;
        }

        internal int Length => _length;
        internal int SourceStart => _sourceStart;
        internal string? MaterializedText => _text;
        internal bool IsString { get; }
        internal bool IsTerminated { get; }
        internal int? EncodedLength { get; }

        internal char GetCharacter<TText>(TText source, int index) where TText : struct, IPdfContentText {
            if ((uint)index >= (uint)_length) throw new ArgumentOutOfRangeException(nameof(index));
            return _text is not null ? _text[index] : source[_sourceStart + index];
        }

        internal string GetText<TText>(TText source, System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText =>
            _text ?? source.Substring(_sourceStart, _length, cancellationToken);

        internal bool Equals<TText>(TText source, string value) where TText : struct, IPdfContentText {
            if (_length != value.Length) return false;
            if (_text is not null) return string.Equals(_text, value, StringComparison.Ordinal);
            for (int index = 0; index < _length; index++) {
                if (source[_sourceStart + index] != value[index]) return false;
            }
            return true;
        }

        internal bool TryParseInt32<TText>(TText source, out int value, System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
            cancellationToken.ThrowIfCancellationRequested();
            if (_length > MaxNumericTokenCharacters) {
                value = default;
                return false;
            }
#if NET8_0_OR_GREATER
            if (_text is null) return source.TryParseInt32(_sourceStart, _length, out value);
            return int.TryParse(_text.AsSpan(), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            if (_text is null && TryParsePlainInt32(source, _sourceStart, _length, out value)) return true;
            return int.TryParse(GetText(source, cancellationToken), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }

        internal bool TryParseDouble<TText>(
            TText source,
            out double value,
            System.Threading.CancellationToken cancellationToken,
            System.Globalization.NumberStyles styles = System.Globalization.NumberStyles.Any) where TText : struct, IPdfContentText {
            cancellationToken.ThrowIfCancellationRequested();
            if (_length > MaxNumericTokenCharacters) {
                value = default;
                return false;
            }
            if (_text is null && (styles & System.Globalization.NumberStyles.Float) == System.Globalization.NumberStyles.Float &&
                TryParsePlainDecimal(source, _sourceStart, _length, out value)) return true;
#if NET8_0_OR_GREATER
            if (_text is null) return source.TryParseDouble(_sourceStart, _length, styles, out value);
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(_text.AsSpan(), styles, System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(GetText(source, cancellationToken), styles, System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }

        /// <summary>
        /// Whether the framework parser accepts this token under <see cref="System.Globalization.NumberStyles.Any"/>,
        /// answered without parsing for the plain <c>[+-]digits[.digits]</c> spelling that is short enough not to overflow.
        /// </summary>
        internal bool IsNumber<TText>(TText source, System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText =>
            (_text is null && IsShortPlainDecimal(source, _sourceStart, _length)) || TryParseDouble(source, out _, cancellationToken);
    }

    private static bool IsShortPlainDecimal<TText>(TText source, int start, int length) where TText : struct, IPdfContentText {
        // 32 characters cannot reach the double range on any target, so the parse cannot fail.
        if (length > 32) return false;
        int index = start;
        int end = start + length;
        if (index < end && (source[index] == '-' || source[index] == '+')) index++;
        bool seenPoint = false;
        bool seenDigit = false;
        for (; index < end; index++) {
            char current = source[index];
            if (current == '.') {
                if (seenPoint) return false;
                seenPoint = true;
            } else if ((uint)(current - '0') <= 9) {
                seenDigit = true;
            } else {
                return false;
            }
        }
        return seenDigit;
    }

#if NET8_0_OR_GREATER
    private static readonly double[] ExactPowersOfTen = {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15,
        1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22
    };

    /// <summary>
    /// Reads <c>[+-]digits[.digits]</c> with a non-zero value of at most 15 significant digits. The digits and the
    /// power of ten are both exact doubles, so one IEEE division gives the correctly rounded value, which is what
    /// the runtime parser returns on these targets. Zero, whose sign the parsers keep, and every other spelling
    /// return false for the runtime parser.
    /// </summary>
    private static bool TryParsePlainDecimal<TText>(TText source, int start, int length, out double value) where TText : struct, IPdfContentText {
        value = 0;
        int index = start;
        int end = start + length;
        bool negative = false;
        if (index < end && (source[index] == '-' || source[index] == '+')) {
            negative = source[index] == '-';
            index++;
        }
        long mantissa = 0;
        int significantDigits = 0;
        int fractionDigits = 0;
        bool seenPoint = false;
        for (; index < end; index++) {
            char current = source[index];
            if (current == '.') {
                if (seenPoint) return false;
                seenPoint = true;
                continue;
            }
            int digit = current - '0';
            if ((uint)digit > 9) return false;
            if (seenPoint) fractionDigits++;
            if (mantissa == 0 && digit == 0) continue;
            if (++significantDigits > 15) return false;
            mantissa = mantissa * 10 + digit;
        }
        if (mantissa == 0 || fractionDigits >= ExactPowersOfTen.Length) return false;
        double magnitude = fractionDigits == 0 ? mantissa : mantissa / ExactPowersOfTen[fractionDigits];
        value = negative ? -magnitude : magnitude;
        return true;
    }
#else
    // Older targets have no span-based number parsing, so every numeric token used to become a substring first.
    // These read the common PDF spellings in place and leave every other spelling to the framework parser.

    /// <summary>Reads <c>[+-]digits</c> that fits an Int32. Any other spelling returns false for the framework parser.</summary>
    private static bool TryParsePlainInt32<TText>(TText source, int start, int length, out int value) where TText : struct, IPdfContentText {
        value = 0;
        int index = start;
        int end = start + length;
        bool negative = false;
        if (index < end && (source[index] == '-' || source[index] == '+')) {
            negative = source[index] == '-';
            index++;
        }
        if (index == end) return false;
        long magnitude = 0;
        for (; index < end; index++) {
            int digit = source[index] - '0';
            if ((uint)digit > 9) return false;
            magnitude = magnitude * 10 + digit;
            if (magnitude > 2147483648L) return false;
        }
        if (negative) magnitude = -magnitude;
        if (magnitude > int.MaxValue) return false;
        value = (int)magnitude;
        return true;
    }

    /// <summary>
    /// Reads a non-zero whole number <c>[+-]digits</c> of at most 15 significant digits, which every target parses
    /// to the same exact double. Fractions are left to the framework parser: .NET Framework does not always round
    /// them correctly, and its own "R" formatting depends on reading back its parser's value. Zero, whose sign the
    /// framework parsers treat differently, and every other spelling also return false.
    /// </summary>
    private static bool TryParsePlainDecimal<TText>(TText source, int start, int length, out double value) where TText : struct, IPdfContentText {
        value = 0;
        int index = start;
        int end = start + length;
        bool negative = false;
        if (index < end && (source[index] == '-' || source[index] == '+')) {
            negative = source[index] == '-';
            index++;
        }
        long mantissa = 0;
        int significantDigits = 0;
        for (; index < end; index++) {
            int digit = source[index] - '0';
            if ((uint)digit > 9) return false;
            if (mantissa == 0 && digit == 0) continue;
            if (++significantDigits > 15) return false;
            mantissa = mantissa * 10 + digit;
        }
        if (mantissa == 0) return false;
        value = negative ? -mantissa : mantissa;
        return true;
    }
#endif

    // PDF dictionaries repeat a small vocabulary across virtually every object. Keep this
    // pool deliberately bounded: arbitrary document names must never be globally interned.
    // Reusing these literals avoids allocating both "/Type" and then "Type" for every
    // catalog, page-tree node, page and resource dictionary parsed by the shared reader.
    private static readonly KeyValuePair<string, string>[] KnownNameTokens = {
        new("/Type", "Type"), new("/Catalog", "Catalog"), new("/Pages", "Pages"), new("/Page", "Page"),
        new("/Parent", "Parent"), new("/MediaBox", "MediaBox"), new("/CropBox", "CropBox"), new("/Rotate", "Rotate"),
        new("/Resources", "Resources"), new("/Contents", "Contents"), new("/Count", "Count"), new("/Kids", "Kids"),
        new("/Root", "Root"), new("/Size", "Size"), new("/Info", "Info"), new("/ID", "ID"),
        new("/Length", "Length"), new("/Filter", "Filter"), new("/FlateDecode", "FlateDecode"), new("/Dests", "Dests"),
        new("/Names", "Names"), new("/Outlines", "Outlines"), new("/OpenAction", "OpenAction"), new("/Metadata", "Metadata"),
        new("/OutputIntents", "OutputIntents"), new("/AcroForm", "AcroForm"), new("/Annots", "Annots"), new("/Subtype", "Subtype"),
        new("/Font", "Font"), new("/XObject", "XObject"), new("/ProcSet", "ProcSet"), new("/Encoding", "Encoding"),
        new("/BaseFont", "BaseFont"), new("/Type1", "Type1"), new("/TrueType", "TrueType"), new("/Widths", "Widths"),
        new("/FirstChar", "FirstChar"), new("/LastChar", "LastChar"), new("/ToUnicode", "ToUnicode"),
        new("/FontDescriptor", "FontDescriptor"), new("/FontFile2", "FontFile2"), new("/FontFile3", "FontFile3"),
        new("/Image", "Image"), new("/Form", "Form"), new("/BBox", "BBox"), new("/Matrix", "Matrix"),
        new("/ColorSpace", "ColorSpace"), new("/DeviceRGB", "DeviceRGB"), new("/DeviceGray", "DeviceGray"),
        new("/DeviceCMYK", "DeviceCMYK"), new("/BitsPerComponent", "BitsPerComponent"), new("/Width", "Width"),
        new("/Height", "Height"), new("/S", "S"), new("/Fit", "Fit"), new("/XYZ", "XYZ"),
        new("/FitH", "FitH"), new("/FitV", "FitV"), new("/FitR", "FitR"), new("/FitB", "FitB"),
        new("/FitBH", "FitBH"), new("/FitBV", "FitBV"), new("/Producer", "Producer"), new("/Creator", "Creator"),
        new("/CreationDate", "CreationDate"), new("/ModDate", "ModDate"), new("/Title", "Title"), new("/Author", "Author"),
        new("/Subject", "Subject"), new("/Keywords", "Keywords"), new("/Trapped", "Trapped")
    };
    private static readonly string[][] KnownTokenTextsByLength = CreateKnownTokenTextsByLength();
    private static readonly string?[][] KnownDecodedNamesByLength = CreateKnownDecodedNamesByLength();

    private static PdfToken MaterializeToken<TText>(TText source, int start, int length,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        if ((uint)length < (uint)KnownTokenTextsByLength.Length && length > 0) {
            string[] candidates = KnownTokenTextsByLength[length];
            // Known names share their leading slash, so candidates are first told apart by the next character.
            int keyIndex = length > 1 ? 1 : 0;
            char key = source[start + keyIndex];
            for (int candidateIndex = 0; candidateIndex < candidates.Length; candidateIndex++) {
                string candidate = candidates[candidateIndex];
                if (candidate[keyIndex] != key) continue;
                bool matches = true;
                for (int characterIndex = 0; characterIndex < length; characterIndex++) {
                    if (source[start + characterIndex] != candidate[characterIndex]) {
                        matches = false;
                        break;
                    }
                }

                if (matches) return new PdfToken(candidate);
            }
        }

        // Numbers are read in place on every target; other tokens are sliced only when their text is needed.
        return new PdfToken(start, length);
    }

    private static string DecodeNameToken<TText>(PdfToken token, TText source,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
        if ((uint)token.Length < (uint)KnownTokenTextsByLength.Length) {
            string[] candidates = KnownTokenTextsByLength[token.Length];
            string?[] decodedNames = KnownDecodedNamesByLength[token.Length];
            for (int i = 0; i < candidates.Length; i++) {
                if (decodedNames[i] is string decoded &&
                    ReferenceEquals(token.MaterializedText, candidates[i])) {
                    return decoded;
                }
            }
        }

        return token.SourceStart >= 0
            ? DecodeName(source, token.SourceStart + 1, token.Length - 1, cancellationToken)
            : DecodeName(new PdfStringContentText(token.MaterializedText!), 1, token.Length - 1, cancellationToken);
    }

    private static bool IsPdfNumberStart(char value) =>
        char.IsDigit(value) || value == '-' || value == '+' || value == '.';

    private static void SetDictionaryItem(PdfDictionary dictionary, string key, PdfObject value) {
#if NET8_0_OR_GREATER
        ref PdfObject? slot = ref CollectionsMarshal.GetValueRefOrAddDefault(dictionary.Items, key, out bool exists);
        dictionary.HasIncompleteSyntax |= exists;
        slot = value;
#else
        dictionary.HasIncompleteSyntax |= dictionary.Items.ContainsKey(key);
        dictionary.Items[key] = value;
#endif
        dictionary.HasIncompleteSyntax |= value.HasIncompleteSyntax;
    }

    private static string[][] CreateKnownTokenTextsByLength() {
        string[] syntaxTokens = { "R", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "true", "false", "null" };
        int maximumLength = 0;
        for (int i = 0; i < syntaxTokens.Length; i++) maximumLength = Math.Max(maximumLength, syntaxTokens[i].Length);
        for (int i = 0; i < KnownNameTokens.Length; i++) maximumLength = Math.Max(maximumLength, KnownNameTokens[i].Key.Length);
        var buckets = new List<string>[maximumLength + 1];
        for (int i = 0; i < syntaxTokens.Length; i++) {
            string token = syntaxTokens[i];
            (buckets[token.Length] ??= new List<string>()).Add(token);
        }
        for (int i = 0; i < KnownNameTokens.Length; i++) {
            string token = KnownNameTokens[i].Key;
            (buckets[token.Length] ??= new List<string>()).Add(token);
        }

        var result = new string[maximumLength + 1][];
        for (int i = 0; i < result.Length; i++) result[i] = buckets[i]?.ToArray() ?? Array.Empty<string>();
        return result;
    }

    private static string?[][] CreateKnownDecodedNamesByLength() {
        var result = new string?[KnownTokenTextsByLength.Length][];
        for (int length = 0; length < result.Length; length++) {
            string[] candidates = KnownTokenTextsByLength[length];
            var decodedNames = new string?[candidates.Length];
            for (int candidateIndex = 0; candidateIndex < candidates.Length; candidateIndex++) {
                string candidate = candidates[candidateIndex];
                for (int nameIndex = 0; nameIndex < KnownNameTokens.Length; nameIndex++) {
                    KeyValuePair<string, string> knownName = KnownNameTokens[nameIndex];
                    if (string.Equals(candidate, knownName.Key, StringComparison.Ordinal)) {
                        decodedNames[candidateIndex] = knownName.Value;
                        break;
                    }
                }
            }
            result[length] = decodedNames;
        }
        return result;
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

    internal static string DecodeName(string raw,
        System.Threading.CancellationToken cancellationToken = default) =>
        DecodeName(raw, 0, raw.Length, cancellationToken);

    private static string DecodeName(string source, int start, int length,
        System.Threading.CancellationToken cancellationToken) =>
        DecodeName(new PdfStringContentText(source), start, length, cancellationToken);

    private static string DecodeName<TText>(TText source, int start, int length,
        System.Threading.CancellationToken cancellationToken) where TText : struct, IPdfContentText {
        cancellationToken.ThrowIfCancellationRequested();
        if (length == 0) return string.Empty;

        bool requiresDecoding = false;
        for (int index = 0; index < length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (source[start + index] == '#' || source[start + index] >= 0x80) {
                requiresDecoding = true;
                break;
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (!requiresDecoding) {
            if (length <= MaxCachedNameLength) return GetCachedName(source, start, length);
            string decoded = source.Substring(start, length, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return decoded;
        }

        var bytes = new byte[length];
        int byteCount = 0;
        bool hasNonAsciiByte = false;
        for (int i = 0; i < length; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char ch = source[start + i];
            if (ch == '#' && i + 2 < length && TryHexNibble(source[start + i + 1], out int hi) && TryHexNibble(source[start + i + 2], out int lo)) {
                byte decoded = (byte)((hi << 4) | lo);
                bytes[byteCount++] = decoded;
                hasNonAsciiByte |= decoded >= 0x80;
                i += 2;
                continue;
            }

            if (ch > byte.MaxValue) {
                cancellationToken.ThrowIfCancellationRequested();
                string fallback = source.Substring(start, length, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                return fallback;
            }
            bytes[byteCount++] = (byte)ch;
            hasNonAsciiByte |= ch >= 0x80;
        }

        if (hasNonAsciiByte) {
            try {
                return PdfEncoding.DecodeCancellable(StrictUtf8NameEncoding, bytes, 0, byteCount, cancellationToken);
            } catch (DecoderFallbackException) {
                // Names are byte sequences. Preserve legacy single-byte names when they are not valid UTF-8.
            }
        }

        return PdfEncoding.Latin1GetStringCancellable(bytes, 0, byteCount, cancellationToken);
    }

    // Plain names repeat across the objects of a document (resource keys, font and tag names).
    // A small per-thread table returns one string per spelling instead of a new one per token.
    private const int MaxCachedNameLength = 48;
    private const int CachedNameSlots = 512;
    [ThreadStatic]
    private static string?[]? t_cachedNames;

    private static string GetCachedName<TText>(TText source, int start, int length) where TText : struct, IPdfContentText {
        uint hash = 2166136261;
        for (int index = 0; index < length; index++) hash = unchecked((hash ^ source[start + index]) * 16777619);
        string?[] names = t_cachedNames ??= new string?[CachedNameSlots];
        int slot = (int)(hash % CachedNameSlots);
        string? cached = names[slot];
        if (cached != null && cached.Length == length) {
            int index = 0;
            while (index < length && cached[index] == source[start + index]) index++;
            if (index == length) return cached;
        }
        string name = source.Substring(start, length);
        names[slot] = name;
        return name;
    }

    private static PdfStringObj CreateParsedString(byte[] bytes, int? encodedTokenLength,
        System.Threading.CancellationToken cancellationToken) {
        string value = PdfTextString.Decode(bytes, cancellationToken);
        return PdfStringObj.FromParsedBytes(
            bytes,
            value,
            // The token can be binary (for example an indexed-image palette). Its decoded
            // view never authorizes changing the original payload during a rewrite.
            useTextStringEncoding: false,
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
