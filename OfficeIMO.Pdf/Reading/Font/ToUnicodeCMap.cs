using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal sealed class ToUnicodeCMap {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private const RegexOptions StaticRegexOptions =
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.NonBacktracking;
#else
    private const RegexOptions StaticRegexOptions =
        RegexOptions.Compiled | RegexOptions.CultureInvariant;
#endif
    // Hex strings can contain whitespace inside source codes and UTF-16 destinations.
    private static readonly Regex BfCharEntryRegex = new Regex(
        @"<(?<src>[0-9A-Fa-f\s\x00]+)>\s*<(?<dst>[0-9A-Fa-f\s\x00]+)>",
        StaticRegexOptions,
        RegexTimeout);
    private static readonly Regex BfCharSectionRegex = new Regex(
        @"beginbfchar([\s\S]*?)endbfchar",
        StaticRegexOptions | RegexOptions.IgnoreCase,
        RegexTimeout);
    private static readonly Regex BfRangeSequentialEntryRegex = new Regex(
        @"<(?<from>[0-9A-Fa-f\s\x00]+)>\s*<(?<to>[0-9A-Fa-f\s\x00]+)>\s*<(?<dst>[0-9A-Fa-f\s\x00]+)>",
        StaticRegexOptions,
        RegexTimeout);
    private static readonly Regex BfRangeSectionRegex = new Regex(
        @"beginbfrange([\s\S]*?)endbfrange",
        StaticRegexOptions | RegexOptions.IgnoreCase,
        RegexTimeout);
    private static readonly Regex BfRangeArrayEntryRegex = new Regex(
        @"<(?<from>[0-9A-Fa-f\s\x00]+)>\s*<(?<to>[0-9A-Fa-f\s\x00]+)>\s*\[(?<dsts>[\s\S]*?)\]",
        StaticRegexOptions | RegexOptions.IgnoreCase,
        RegexTimeout);

    private const int MaxMappings = 65536;
    private const int MaxRangeMappings = 4096;
    private const int MaxSourceCodeBytes = 4;
    private const int MaxDestinationTextCharacters = 4096;
    private const int MaxReverseMappingTextCharacters = 64;
    private const int MaxReverseMapNodes = 65536;
    private const int MaxDestinationHexCharactersWithWhitespace = MaxDestinationTextCharacters * 8;

    private readonly Dictionary<string, string> _map = new(StringComparer.OrdinalIgnoreCase);
    private readonly ReverseMapNode _reverseMapRoot = new();
    private int _maxKeyBytes = 1;
    private int _processedMappings;
    private int _reverseMapNodeCount = 1;
    private bool _reverseMapBudgetExhausted;
    private long _retainedMappingTextCharacters;
    private long _retainedReverseCodeCharacters;
    private int _reverseMapTerminalCount;

    internal int MappingCount => _map.Count;

    /// <summary>Conservative retained-size estimate used when admitting a font to the document cache.</summary>
    /// <remarks>Allowances cover mapping entries and strings, plus reverse nodes, sparse child dictionaries
    /// and their capacity. Reverse terminal strings can outlive replacement of a forward mapping.</remarks>
    internal long EstimatedRetainedBytes =>
        _map.Count * 128L + _retainedMappingTextCharacters * sizeof(char) +
        _reverseMapNodeCount * 256L + _reverseMapTerminalCount * 24L +
        _retainedReverseCodeCharacters * sizeof(char);

    /// <summary>Source code hex strings and their mapped text.</summary>
    internal IEnumerable<KeyValuePair<string, string>> Mappings => _map;

    public static bool TryParse(byte[] data, out ToUnicodeCMap? cmap) {
        try {
            string s = PdfEncoding.Latin1GetString(data);
            var inst = new ToUnicodeCMap();
            inst.Parse(s);
            cmap = inst; return true;
        } catch { cmap = null; return false; }
    }

    private void Parse(string s) {
        // Handle beginbfchar / endbfchar lines like: <AB> <00E9>
        foreach (Match section in BfCharSectionRegex.Matches(s)) {
            foreach (Match m in BfCharEntryRegex.Matches(section.Groups[1].Value)) {
                if (HasReachedMappingLimit()) {
                    break;
                }

                AddMap(m.Groups["src"].Value, m.Groups["dst"].Value);
            }

            if (HasReachedMappingLimit()) {
                break;
            }
        }
        // Handle beginbfrange / endbfrange, sequential mapping
        foreach (Match section in BfRangeSectionRegex.Matches(s)) {
            if (HasReachedMappingLimit()) {
                break;
            }

            string body = section.Groups[1].Value;
            string sequentialBody = BfRangeArrayEntryRegex.Replace(body, string.Empty);
            foreach (Match m in BfRangeSequentialEntryRegex.Matches(sequentialBody)) {
                if (HasReachedMappingLimit()) {
                    break;
                }

                string fromHex = RemoveHexWhitespace(m.Groups["from"].Value);
                uint from = Convert.ToUInt32(fromHex, 16);
                uint to = Convert.ToUInt32(RemoveHexWhitespace(m.Groups["to"].Value), 16);
                uint dst = Convert.ToUInt32(RemoveHexWhitespace(m.Groups["dst"].Value), 16);
                ulong rangeLength = to >= from ? (ulong)to - from + 1 : 0;
                if (rangeLength == 0 || rangeLength > MaxRangeMappings || rangeLength > (ulong)(MaxMappings - _processedMappings)) {
                    continue;
                }

                int keyBytes = fromHex.Length / 2; // bytes per code
                // Bound iterations by the validated count, including a range ending at FFFFFFFF.
                for (int index = 0; index < (int)rangeLength; index++) {
                    uint code = from + (uint)index;
                    uint u = unchecked(dst + (uint)index);
                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    string dstHex = u.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
                    AddMap(srcHex, dstHex);
                }
            }

            foreach (Match m in BfRangeArrayEntryRegex.Matches(body)) {
                if (HasReachedMappingLimit()) {
                    break;
                }

                string fromHex = RemoveHexWhitespace(m.Groups["from"].Value);
                uint from = Convert.ToUInt32(fromHex, 16);
                uint to = Convert.ToUInt32(RemoveHexWhitespace(m.Groups["to"].Value), 16);
                if (to < from) {
                    continue;
                }

                int rangeLength = (int)Math.Min((ulong)MaxRangeMappings, (ulong)to - from + 1);
                int keyBytes = fromHex.Length / 2;
                int index = 0;
                foreach (string destination in ReadHexArrayEntries(m.Groups["dsts"].Value)) {
                    if (index >= rangeLength || HasReachedMappingLimit()) {
                        break;
                    }

                    uint code = from + (uint)index;
                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    AddMap(srcHex, destination);
                    index++;
                }
            }
        }
    }

    private void AddMap(string srcHex, string dstHex) {
        if (HasReachedMappingLimit()) {
            return;
        }

        _processedMappings++;

        if (dstHex.Length > MaxDestinationHexCharactersWithWhitespace) {
            return;
        }

        srcHex = RemoveHexWhitespace(srcHex);
        dstHex = RemoveHexWhitespace(dstHex);
        if (srcHex.Length % 2 != 0) srcHex = "0" + srcHex;
        if (srcHex.Length == 0 || srcHex.Length / 2 > MaxSourceCodeBytes || dstHex.Length > MaxDestinationTextCharacters * 4) {
            return;
        }

        _maxKeyBytes = Math.Max(_maxKeyBytes, srcHex.Length / 2);
        string key = srcHex.ToUpperInvariant();
        // dst may be multi-codepoints; keep as UTF-16 string
        string s = HexToString(dstHex);
        if (_map.TryGetValue(key, out string? previous)) _retainedMappingTextCharacters -= previous.Length;
        else _retainedMappingTextCharacters += key.Length;
        _retainedMappingTextCharacters += s.Length;
        _map[key] = s;
        if (!_reverseMapBudgetExhausted && s.Length > 0 && s.Length <= MaxReverseMappingTextCharacters) {
            ReverseMapNode node = _reverseMapRoot;
            for (int index = 0; index < s.Length; index++) {
                if (!node.TryGetOrAdd(s[index], ref _reverseMapNodeCount, MaxReverseMapNodes, out node)) {
                    _reverseMapBudgetExhausted = true;
                    return;
                }
            }

            if (node.CodeHex is null) {
                node.CodeHex = key;
                // Reverse terminals and partially created branches survive forward-map replacement.
                // Charge their source strings separately, including conservative string-object overhead.
                _retainedReverseCodeCharacters += key.Length;
                _reverseMapTerminalCount++;
            }
        }
    }

    private bool HasReachedMappingLimit() => _processedMappings >= MaxMappings;

    private static string RemoveHexWhitespace(string value) {
        bool hasWhitespace = false;
        for (int i = 0; i < value.Length; i++) {
            if (char.IsWhiteSpace(value[i]) || value[i] == '\0') {
                hasWhitespace = true;
                break;
            }
        }

        if (!hasWhitespace) {
            return value;
        }

        var sb = new System.Text.StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i]) && value[i] != '\0') {
                sb.Append(value[i]);
            }
        }

        return sb.ToString();
    }

    private static IEnumerable<string> ReadHexArrayEntries(string body) {
        int index = 0;
        while (index < body.Length) {
            int start = body.IndexOf('<', index);
            if (start < 0) {
                yield break;
            }

            int end = body.IndexOf('>', start + 1);
            if (end < 0) {
                yield break;
            }

            yield return body.Substring(start + 1, end - start - 1);
            index = end + 1;
        }
    }

    private static string HexToString(string hex) {
        // Interpret as sequence of 16-bit big-endian code points
        if (hex.Length % 4 != 0) hex = hex.PadLeft(((hex.Length + 3) / 4) * 4, '0');
        var chars = new List<char>(hex.Length / 4);
        for (int i = 0; i < hex.Length; i += 4) {
            ushort code = Convert.ToUInt16(hex.Substring(i, 4), 16);
            chars.Add((char)code);
        }
        return new string(chars.ToArray());
    }

    public string MapBytes(byte[] bytes) => MapBytes(bytes, PdfReadLimits.DefaultMaxDecodedTextCharacters);

    internal string MapBytes(byte[] bytes, int maxOutputCharacters) =>
        MapBytesCore(bytes, maxOutputCharacters, allowUnmappedBytes: true, out _);

    /// <summary>Maps every supplied code without falling back to raw bytes.</summary>
    internal bool TryMapBytes(byte[] bytes, int maxOutputCharacters, out string decoded) {
        decoded = MapBytesCore(bytes, maxOutputCharacters, allowUnmappedBytes: false, out bool allMapped);
        return allMapped;
    }

    private string MapBytesCore(byte[] bytes, int maxOutputCharacters, bool allowUnmappedBytes, out bool allMapped) {
        if (maxOutputCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputCharacters), maxOutputCharacters, "Maximum decoded text characters must be positive.");
        }

        allMapped = true;
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < bytes.Length;) {
            // Greedy match up to _maxKeyBytes (1-2 bytes typical)
            int max = Math.Min(_maxKeyBytes, bytes.Length - i);
            string? mapped = null; int used = 0;
            for (int len = max; len >= 1; len--) {
                string key = ByteSliceToHex(bytes, i, len);
                if (_map.TryGetValue(key, out var s)) { mapped = s; used = len; break; }
            }
            if (mapped is null) {
                allMapped = false;
                if (!allowUnmappedBytes) return string.Empty;
            }
            int appendLength = mapped?.Length ?? 1;
            long nextLength = (long)sb.Length + appendLength;
            if (nextLength > maxOutputCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, nextLength);
            }

            if (mapped is null) { sb.Append((char)bytes[i]); i++; } else { sb.Append(mapped); i += used; }
        }
        return sb.ToString();
    }

    public bool TryEncodeText(string text, out string hex) {
        Guard.NotNull(text, nameof(text));

        var sb = new System.Text.StringBuilder(text.Length * 4);
        if (!TryEncodeTextCodes(text, out IReadOnlyList<string> codeHexValues)) {
            hex = string.Empty;
            return false;
        }

        foreach (string codeHex in codeHexValues) {
            sb.Append(codeHex);
        }

        hex = sb.ToString();
        return true;
    }

    internal bool TryEncodeTextCodes(string text, out IReadOnlyList<string> codeHexValues) {
        Guard.NotNull(text, nameof(text));

        var codes = new List<string>();
        for (int index = 0; index < text.Length;) {
            string? codeHex = null;
            int matchedLength = 0;
            ReverseMapNode node = _reverseMapRoot;
            int maxLength = Math.Min(MaxReverseMappingTextCharacters, text.Length - index);
            for (int length = 1; length <= maxLength; length++) {
                if (!node.TryGet(text[index + length - 1], out node)) {
                    break;
                }

                if (node.CodeHex != null) {
                    codeHex = node.CodeHex;
                    matchedLength = length;
                }
            }

            if (codeHex == null) {
                codeHexValues = Array.Empty<string>();
                return false;
            }

            codes.Add(codeHex);
            index += matchedLength;
        }

        codeHexValues = codes;
        return true;
    }

    private sealed class ReverseMapNode {
        private Dictionary<char, ReverseMapNode>? _children;

        internal string? CodeHex { get; set; }

        internal bool TryGetOrAdd(char value, ref int nodeCount, int maximumNodes, out ReverseMapNode child) {
            if (_children != null && _children.TryGetValue(value, out ReverseMapNode? existing)) {
                child = existing;
                return true;
            }

            if (nodeCount >= maximumNodes) {
                child = null!;
                return false;
            }

            _children ??= new Dictionary<char, ReverseMapNode>();
            child = new ReverseMapNode();
            _children[value] = child;
            nodeCount++;
            return true;
        }

        internal bool TryGet(char value, out ReverseMapNode child) {
            if (_children != null && _children.TryGetValue(value, out ReverseMapNode? match)) {
                child = match;
                return true;
            }

            child = null!;
            return false;
        }
    }

    private static string ByteSliceToHex(byte[] b, int start, int len) {
        var sb = new System.Text.StringBuilder(len * 2);
        for (int i = 0; i < len; i++) sb.Append(b[start + i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        return sb.ToString();
    }

}
