#if NET8_0_OR_GREATER
using System.Buffers;
using System.Runtime.InteropServices;
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
            PdfToken token = tokens[i];
            if (i == 0 && token.Equals(tokens.Source, "<<")) continue;
            if (token.Length > 0 && token.GetCharacter(tokens.Source, 0) == '/') {
                string key = DecodeNameToken(token, tokens.Source);
                if (i + 1 < tokens.Count && !tokens[i + 1].Equals(tokens.Source, ">>")) {
                    var (obj, consumed) = ParseObject(tokens, i + 1, effectiveLimits, 0);
                    SetDictionaryItem(d, key, obj);
                    i += consumed + 1;
                } else d.HasIncompleteSyntax = true;
            } else if (token.Equals(tokens.Source, ">>") && i == tokens.Count - 1) {
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
        string source = tokens.Source;
        if (token.Equals(source, "<<")) {
            var dict = new PdfDictionary();
            int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, ">>")) {
                PdfToken keyToken = tokens[j];
                if (keyToken.Length > 0 && keyToken.GetCharacter(source, 0) == '/') {
                    string key = DecodeNameToken(keyToken, source);
                    if (j + 1 < tokens.Count && !tokens[j + 1].Equals(source, ">>")) {
                        var (obj, consumed) = ParseObject(tokens, j + 1, limits, depth + 1);
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
            var arr = new PdfArray(EstimateArrayItemCount(tokens, i)); int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, "]")) {
                var (inner, used) = ParseObject(tokens, j, limits, depth + 1);
                arr.Items.Add(inner);
                arr.HasIncompleteSyntax |= inner.HasIncompleteSyntax;
                j += used + 1;
            }
            arr.HasIncompleteSyntax |= j >= tokens.Count;
            return (arr, j - i);
        }
        if (token.Length > 0 && token.GetCharacter(source, 0) == '/') return (new PdfName(DecodeNameToken(token, source)), 0);
        if (token.IsString && token.Length > 0 && token.GetCharacter(source, 0) == '(') {
            bool isTerminated = token.IsTerminated;
            int innerLength = token.Length - (isTerminated ? 2 : 1);
            var value = CreateParsedString(
                PdfStringParser.ParseLiteralToBytes(source, token.SourceStart + 1, innerLength),
                token.EncodedLength);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (token.IsString && token.Length > 0 && token.GetCharacter(source, 0) == '<' &&
            (token.Length == 1 || token.GetCharacter(source, 1) != '<')) {
            bool isTerminated = token.IsTerminated;
            int innerLength = token.Length - (isTerminated ? 2 : 1);
            var value = CreateParsedString(
                PdfTextString.DecodeHexBytes(source, token.SourceStart + 1, innerLength),
                token.EncodedLength);
            value.HasIncompleteSyntax = !isTerminated;
            return (value, 0);
        }
        if (token.Equals(source, "true")) return (new PdfBoolean(true), 0);
        if (token.Equals(source, "false")) return (new PdfBoolean(false), 0);
        if (token.Equals(source, "null")) return (PdfNull.Instance, 0);
        if (token.Length > 0 && IsPdfNumberStart(token.GetCharacter(source, 0))) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2].Equals(source, "R") &&
                token.TryParseInt32(source, out int obj) && tokens[i + 1].TryParseInt32(source, out int gen)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (token.TryParseDouble(source, out double val)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(token.GetText(source)) { HasIncompleteSyntax = true }, 0);
    }

    private static int EstimateArrayItemCount(in PooledTokenBuffer tokens, int arrayStart) {
        int count = 0;
        int index = arrayStart + 1;
        string source = tokens.Source;
        while (index < tokens.Count) {
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
                tokens[index].TryParseInt32(source, out _) &&
                tokens[index + 1].TryParseInt32(source, out _)) {
                index += 3;
                continue;
            }

            index++;
        }

        return count;
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
        var tokens = new PooledTokenBuffer(s, Math.Min(estimatedTokens, effectiveLimits.MaxTokensPerObject));
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
            if (c == '[' || c == ']') { tokens.Add(new PdfToken(c == '[' ? "[" : "]")); i++; continue; }
            if (c == '<') {
                int tokenStart = i++;
                while (i < end && s[i] != '>') i++;
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
            while (j < end && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
            int tokenLength = j - i;
            if (tokenLength == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < end && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tokenLength = j - i;
            }
            if (tokenLength == 0) {
                // A malformed or unexpected standalone delimiter must still consume input.
                // Otherwise ')' and '>' repeatedly produce empty tokens until the token
                // budget is exhausted, turning one bad byte into excessive CPU and memory use.
                tokenLength = 1;
                j = i + 1;
            }
            tokens.Add(MaterializeToken(s, i, tokenLength));
            if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
            }
            i = j;
        }

        if (tokens.Count > effectiveLimits.MaxTokensPerObject) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectTokens, effectiveLimits.MaxTokensPerObject, tokens.Count);
        }
    }

    /// <summary>Reuses token arrays on modern targets while retaining the compatible list path on legacy targets.</summary>
    private struct PooledTokenBuffer : IDisposable {
#if NET8_0_OR_GREATER
        private PdfToken[]? _rented;
#else
        private List<PdfToken>? _small;
#endif
        private int _count;
        private readonly string _source;

        internal PooledTokenBuffer(string source, int initialCapacity) {
            _source = source;
#if NET8_0_OR_GREATER
            _rented = ArrayPool<PdfToken>.Shared.Rent(Math.Max(initialCapacity, 16));
#else
            _small = new List<PdfToken>(initialCapacity);
#endif
        }

        internal readonly int Count => _count;
        internal readonly string Source => _source;

        internal readonly PdfToken this[int index] {
            get {
                if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index));
#if NET8_0_OR_GREATER
                return _rented![index];
#else
                return _small![index];
#endif
            }
        }

        internal void Add(PdfToken token) {
#if NET8_0_OR_GREATER
            if (_count == _rented!.Length) {
                PdfToken[] expanded = ArrayPool<PdfToken>.Shared.Rent(checked(_count * 2));
                Array.Copy(_rented, expanded, _count);
                ArrayPool<PdfToken>.Shared.Return(_rented, clearArray: true);
                _rented = expanded;
            }

            _rented[_count] = token;
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

        internal char GetCharacter(string source, int index) {
            if ((uint)index >= (uint)_length) throw new ArgumentOutOfRangeException(nameof(index));
            return _text is not null ? _text[index] : source[_sourceStart + index];
        }

        internal string GetText(string source) => _text ?? source.Substring(_sourceStart, _length);

        internal bool Equals(string source, string value) {
            if (_length != value.Length) return false;
            if (_text is not null) return string.Equals(_text, value, StringComparison.Ordinal);
            return string.CompareOrdinal(source, _sourceStart, value, 0, _length) == 0;
        }

        internal bool TryParseInt32(string source, out int value) {
#if NET8_0_OR_GREATER
            ReadOnlySpan<char> span = _text is not null ? _text.AsSpan() : source.AsSpan(_sourceStart, _length);
            return int.TryParse(span, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            return int.TryParse(GetText(source), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }

        internal bool TryParseDouble(
            string source,
            out double value,
            System.Globalization.NumberStyles styles = System.Globalization.NumberStyles.Any) {
#if NET8_0_OR_GREATER
            ReadOnlySpan<char> span = _text is not null ? _text.AsSpan() : source.AsSpan(_sourceStart, _length);
            return double.TryParse(span, styles,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            return double.TryParse(GetText(source), styles,
                System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }
    }

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

    private static PdfToken MaterializeToken(string source, int start, int length) {
        if ((uint)length < (uint)KnownTokenTextsByLength.Length) {
            string[] candidates = KnownTokenTextsByLength[length];
            for (int candidateIndex = 0; candidateIndex < candidates.Length; candidateIndex++) {
                string candidate = candidates[candidateIndex];
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

#if NET8_0_OR_GREATER
        return new PdfToken(start, length);
#else
        // Older targets do not expose span-based numeric parsing. Preserve their
        // single materialization rather than recreating the same slice at each parse.
        return new PdfToken(source.Substring(start, length));
#endif
    }

    private static string DecodeNameToken(PdfToken token, string source) {
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

        string encodedName = token.SourceStart >= 0
            ? source.Substring(token.SourceStart + 1, token.Length - 1)
            : token.GetText(source).Substring(1);
        return DecodeName(encodedName);
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
