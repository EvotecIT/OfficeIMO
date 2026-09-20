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
            string tokenText = tokens[i].Text;
            if (i == 0 && tokenText == "<<") continue;
            if (tokenText.Length > 0 && tokenText[0] == '/') {
                string key = DecodeNameToken(tokenText);
                if (i + 1 < tokens.Count && tokens[i + 1].Text != ">>") {
                    var (obj, consumed) = ParseObject(tokens, i + 1, effectiveLimits, 0);
                    SetDictionaryItem(d, key, obj);
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
                    string key = DecodeNameToken(keyToken);
                    if (j + 1 < tokens.Count && tokens[j + 1].Text != ">>") {
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
        if (tok == "[") {
            var arr = new PdfArray(EstimateArrayItemCount(tokens, i)); int j = i + 1;
            while (j < tokens.Count && tokens[j].Text != "]") {
                var (inner, used) = ParseObject(tokens, j, limits, depth + 1);
                arr.Items.Add(inner);
                arr.HasIncompleteSyntax |= inner.HasIncompleteSyntax;
                j += used + 1;
            }
            arr.HasIncompleteSyntax |= j >= tokens.Count;
            return (arr, j - i);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(DecodeNameToken(tok)), 0);
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

    private static int EstimateArrayItemCount(in PooledTokenBuffer tokens, int arrayStart) {
        int count = 0;
        int index = arrayStart + 1;
        while (index < tokens.Count) {
            string token = tokens[index].Text;
            if (token == "]") break;
            count++;
            if (token == "[" || token == "<<") {
                // Do not rescan nested subtrees. Their own ParseObject call can size
                // a flat child array, while the parent grows from observed entries.
                return 0;
            }

            if (index + 2 < tokens.Count &&
                tokens[index + 2].Text == "R" &&
                int.TryParse(tokens[index].Text, out _) &&
                int.TryParse(tokens[index + 1].Text, out _)) {
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
            if (c == '[' || c == ']') { tokens.Add(new PdfToken(c == '[' ? "[" : "]")); i++; continue; }
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
            string tok = MaterializeToken(s, i, j - i);
            if (tok.Length == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < end && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tok = MaterializeToken(s, i, j - i);
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

    private static string MaterializeToken(string source, int start, int length) {
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

                if (matches) return candidate;
            }
        }

        return source.Substring(start, length);
    }

    private static string DecodeNameToken(string token) {
        if ((uint)token.Length < (uint)KnownTokenTextsByLength.Length) {
            string[] candidates = KnownTokenTextsByLength[token.Length];
            string?[] decodedNames = KnownDecodedNamesByLength[token.Length];
            for (int i = 0; i < candidates.Length; i++) {
                if (decodedNames[i] is string decoded && ReferenceEquals(token, candidates[i])) {
                    return decoded;
                }
            }
        }

        return DecodeName(token.Substring(1));
    }

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
