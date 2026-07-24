using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal sealed class ToUnicodeCMap {
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
        var bfchar = new Regex(@"<(?<src>[0-9A-Fa-f]+)>\s+<(?<dst>[0-9A-Fa-f]+)>");
        foreach (Match section in Regex.Matches(s, @"beginbfchar([\s\S]*?)endbfchar", RegexOptions.IgnoreCase)) {
            foreach (Match m in bfchar.Matches(section.Groups[1].Value)) {
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
        var bfrangeLine = new Regex(@"<(?<from>[0-9A-Fa-f]+)>\s+<(?<to>[0-9A-Fa-f]+)>\s+<(?<dst>[0-9A-Fa-f]+)>");
        foreach (Match section in Regex.Matches(s, @"beginbfrange([\s\S]*?)endbfrange", RegexOptions.IgnoreCase)) {
            if (HasReachedMappingLimit()) {
                break;
            }

            string body = section.Groups[1].Value;
            string sequentialBody = Regex.Replace(body, @"<(?<from>[0-9A-Fa-f]+)>\s+<(?<to>[0-9A-Fa-f]+)>\s+\[(?<dsts>[\s\S]*?)\]", string.Empty, RegexOptions.IgnoreCase);
            foreach (Match m in bfrangeLine.Matches(sequentialBody)) {
                if (HasReachedMappingLimit()) {
                    break;
                }

                int from = Convert.ToInt32(m.Groups["from"].Value, 16);
                int to = Convert.ToInt32(m.Groups["to"].Value, 16);
                int dst = Convert.ToInt32(m.Groups["dst"].Value, 16);
                int rangeLength = to >= from ? to - from + 1 : 0;
                if (rangeLength <= 0 || rangeLength > MaxRangeMappings || _processedMappings + rangeLength > MaxMappings) {
                    continue;
                }

                int keyBytes = m.Groups["from"].Value.Length / 2; // bytes per code
                for (int code = from, u = dst; code <= to; code++, u++) {
                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    string dstHex = u.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
                    AddMap(srcHex, dstHex);
                }
            }

            foreach (Match m in Regex.Matches(body, @"<(?<from>[0-9A-Fa-f]+)>\s+<(?<to>[0-9A-Fa-f]+)>\s+\[(?<dsts>[\s\S]*?)\]", RegexOptions.IgnoreCase)) {
                if (HasReachedMappingLimit()) {
                    break;
                }

                int from = Convert.ToInt32(m.Groups["from"].Value, 16);
                int to = Convert.ToInt32(m.Groups["to"].Value, 16);
                int keyBytes = m.Groups["from"].Value.Length / 2;
                int code = from;
                foreach (string destination in ReadHexArrayEntries(m.Groups["dsts"].Value)) {
                    if (code > to) {
                        break;
                    }

                    if (HasReachedMappingLimit() || code - from >= MaxRangeMappings) {
                        break;
                    }

                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    AddMap(srcHex, destination);
                    code++;
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
        _map[key] = s;
        if (!_reverseMapBudgetExhausted && s.Length > 0 && s.Length <= MaxReverseMappingTextCharacters) {
            ReverseMapNode node = _reverseMapRoot;
            for (int index = 0; index < s.Length; index++) {
                if (!node.TryGetOrAdd(s[index], ref _reverseMapNodeCount, MaxReverseMapNodes, out node)) {
                    _reverseMapBudgetExhausted = true;
                    return;
                }
            }

            node.CodeHex ??= key;
        }
    }

    private bool HasReachedMappingLimit() => _processedMappings >= MaxMappings;

    private static string RemoveHexWhitespace(string value) {
        bool hasWhitespace = false;
        for (int i = 0; i < value.Length; i++) {
            if (char.IsWhiteSpace(value[i])) {
                hasWhitespace = true;
                break;
            }
        }

        if (!hasWhitespace) {
            return value;
        }

        var sb = new System.Text.StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) {
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

    internal string MapBytes(byte[] bytes, int maxOutputCharacters) {
        if (maxOutputCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputCharacters), maxOutputCharacters, "Maximum decoded text characters must be positive.");
        }

        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < bytes.Length;) {
            // Greedy match up to _maxKeyBytes (1-2 bytes typical)
            int max = Math.Min(_maxKeyBytes, bytes.Length - i);
            string? mapped = null; int used = 0;
            for (int len = max; len >= 1; len--) {
                string key = ByteSliceToHex(bytes, i, len);
                if (_map.TryGetValue(key, out var s)) { mapped = s; used = len; break; }
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
