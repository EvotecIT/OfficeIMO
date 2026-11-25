namespace OfficeIMO.Pdf;

internal sealed class ToUnicodeCMap {
    private const int MaxMappings = 65536;
    private const int MaxRangeMappings = 4096;
    private const int MaxSourceCodeBytes = 4;
    private const int MaxDestinationTextCharacters = 4096;
    private const int MaxReverseMappingTextCharacters = 64;
    private const int MaxReverseMapNodes = 65536;
    private const int MaxDestinationHexCharactersWithWhitespace = MaxDestinationTextCharacters * 8;

    // Keyed by source-code byte count and value so lookups never format the code bytes as hex text.
    // A balanced tree bounds lookup cost for untrusted codes, including colliding integer hashes.
    private readonly SortedDictionary<ulong, string> _map = new();
    // Text is decoded one glyph at a time, so lookups dominate. One- and two-byte codes, the common
    // case, are mirrored in tables with bounded lookups: a direct array and an integer-keyed dictionary
    // whose 65,536 possible keys cannot be made to share a hash bucket chain beyond its prime modulus.
    private string?[]? _oneByteMap;
    private Dictionary<int, string>? _twoByteMap;
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
    internal IEnumerable<KeyValuePair<string, string>> Mappings {
        get {
            foreach (KeyValuePair<ulong, string> mapping in _map) {
                yield return new KeyValuePair<string, string>(FormatKey(mapping.Key), mapping.Value);
            }
        }
    }

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
        int position = 0;
        while (TryFindSection(s, "beginbfchar", "endbfchar", ref position, out int bodyStart, out int bodyEnd)) {
            int entry = bodyStart;
            while (!HasReachedMappingLimit() && TryMatchHexEntry(s, ref entry, bodyEnd, 2, out HexGroups groups)) {
                AddMap(groups.Value(s, 0), groups.Value(s, 1));
            }

            if (HasReachedMappingLimit()) {
                break;
            }
        }
        // Handle beginbfrange / endbfrange, sequential mapping
        position = 0;
        while (TryFindSection(s, "beginbfrange", "endbfrange", ref position, out int bodyStart, out int bodyEnd)) {
            if (HasReachedMappingLimit()) {
                break;
            }

            string body = s.Substring(bodyStart, bodyEnd - bodyStart);
            string sequentialBody = RemoveRangeArrayEntries(body);
            int entry = 0;
            while (!HasReachedMappingLimit() && TryMatchHexEntry(sequentialBody, ref entry, sequentialBody.Length, 3, out HexGroups groups)) {
                string fromHex = RemoveHexWhitespace(groups.Value(sequentialBody, 0));
                uint from = Convert.ToUInt32(fromHex, 16);
                uint to = Convert.ToUInt32(RemoveHexWhitespace(groups.Value(sequentialBody, 1)), 16);
                string dstHex = RemoveHexWhitespace(groups.Value(sequentialBody, 2));
                ulong rangeLength = to >= from ? (ulong)to - from + 1 : 0;
                if (rangeLength == 0 || rangeLength > MaxRangeMappings || rangeLength > (ulong)(MaxMappings - _processedMappings)) {
                    continue;
                }

                int keyBytes = fromHex.Length / 2; // bytes per code
                // Bound iterations by the validated count, including a range ending at FFFFFFFF.
                for (int index = 0; index < (int)rangeLength; index++) {
                    uint code = from + (uint)index;
                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    AddMap(srcHex, OffsetRangeDestination(dstHex, (uint)index));
                }
            }

            entry = 0;
            while (!HasReachedMappingLimit() && TryMatchRangeArrayEntry(body, ref entry, out HexGroups groups, out int destinationsStart, out int destinationsEnd)) {
                string fromHex = RemoveHexWhitespace(groups.Value(body, 0));
                uint from = Convert.ToUInt32(fromHex, 16);
                uint to = Convert.ToUInt32(RemoveHexWhitespace(groups.Value(body, 1)), 16);
                if (to < from) {
                    continue;
                }

                int rangeLength = (int)Math.Min((ulong)MaxRangeMappings, (ulong)to - from + 1);
                int keyBytes = fromHex.Length / 2;
                int index = 0;
                foreach (string destination in ReadHexArrayEntries(body.Substring(destinationsStart, destinationsEnd - destinationsStart))) {
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

    // The CMap is scanned directly instead of with regular expressions. Each scan reproduces the matches
    // of the patterns it replaced: sections are "begin...end" keywords matched without regard to case
    // with the shortest body, and entries are hex strings of [0-9A-Fa-f], whitespace or NUL characters.

    /// <summary>Character offsets of up to three hex-string groups of one entry.</summary>
    private struct HexGroups {
        internal int Start0, End0, Start1, End1, Start2, End2;

        internal readonly string Value(string text, int group) => group switch {
            0 => text.Substring(Start0, End0 - Start0),
            1 => text.Substring(Start1, End1 - Start1),
            _ => text.Substring(Start2, End2 - Start2)
        };

        internal void Set(int group, int start, int end) {
            switch (group) {
                case 0: Start0 = start; End0 = end; break;
                case 1: Start1 = start; End1 = end; break;
                default: Start2 = start; End2 = end; break;
            }
        }
    }

    private static bool TryFindSection(string text, string begin, string end, ref int position, out int bodyStart, out int bodyEnd) {
        bodyStart = bodyEnd = 0;
        if (position >= text.Length) return false;
        int beginIndex = text.IndexOf(begin, position, StringComparison.OrdinalIgnoreCase);
        if (beginIndex < 0) return false;
        int endIndex = text.IndexOf(end, beginIndex + begin.Length, StringComparison.OrdinalIgnoreCase);
        // Without a closing keyword after this one there is none after any later opening keyword either.
        if (endIndex < 0) return false;
        bodyStart = beginIndex + begin.Length;
        bodyEnd = endIndex;
        position = endIndex + end.Length;
        return true;
    }

    private static bool IsHexEntryCharacter(char value) =>
        value is >= '0' and <= '9' or >= 'A' and <= 'F' or >= 'a' and <= 'f' or '\0' || char.IsWhiteSpace(value);

    /// <summary>Matches consecutive hex strings, separated only by whitespace, starting at <paramref name="start"/>.</summary>
    private static int MatchHexGroupsAt(string text, int start, int limit, int groupCount, ref HexGroups groups) {
        int index = start;
        for (int group = 0; group < groupCount; group++) {
            if (group > 0) {
                while (index < limit && char.IsWhiteSpace(text[index])) index++;
            }
            if (index >= limit || text[index] != '<') return -1;
            int valueStart = ++index;
            while (index < limit && IsHexEntryCharacter(text[index])) index++;
            if (index == valueStart || index >= limit || text[index] != '>') return -1;
            groups.Set(group, valueStart, index);
            index++;
        }
        return index;
    }

    /// <summary>Finds the next entry of <paramref name="groupCount"/> hex strings at or after <paramref name="position"/>.</summary>
    private static bool TryMatchHexEntry(string text, ref int position, int limit, int groupCount, out HexGroups groups) {
        groups = default;
        while (position < limit) {
            int start = text.IndexOf('<', position, limit - position);
            if (start < 0) {
                position = limit;
                return false;
            }
            int end = MatchHexGroupsAt(text, start, limit, groupCount, ref groups);
            if (end >= 0) {
                position = end;
                return true;
            }
            position = start + 1;
        }
        return false;
    }

    /// <summary>Finds the next "&lt;from&gt; &lt;to&gt; [destinations]" range entry; the destinations end at the first ']'.</summary>
    private static bool TryMatchRangeArrayEntry(string text, ref int position, out HexGroups groups,
        out int destinationsStart, out int destinationsEnd) {
        groups = default;
        destinationsStart = destinationsEnd = 0;
        int nextClose = -1;
        while (position < text.Length) {
            int start = text.IndexOf('<', position);
            if (start < 0) break;
            int index = MatchHexGroupsAt(text, start, text.Length, 2, ref groups);
            if (index >= 0) {
                while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
                if (index < text.Length && text[index] == '[') {
                    if (nextClose <= index) nextClose = text.IndexOf(']', index + 1);
                    // With no ']' left, no later entry can close either.
                    if (nextClose < 0) break;
                    destinationsStart = index + 1;
                    destinationsEnd = nextClose;
                    position = nextClose + 1;
                    return true;
                }
            }
            position = start + 1;
        }
        position = text.Length;
        return false;
    }

    private static string RemoveRangeArrayEntries(string body) {
        int position = 0, copied = 0;
        StringBuilder? remaining = null;
        while (TryMatchRangeArrayEntry(body, ref position, out HexGroups groups, out _, out _)) {
            // The entry begins at its first hex string, one character before the group.
            int entryStart = groups.Start0 - 1;
            remaining ??= new StringBuilder(body.Length);
            remaining.Append(body, copied, entryStart - copied);
            copied = position;
        }
        if (remaining is null) return body;
        remaining.Append(body, copied, body.Length - copied);
        return remaining.ToString();
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

        int keyBytes = srcHex.Length / 2;
        _maxKeyBytes = Math.Max(_maxKeyBytes, keyBytes);
        ulong key = CreateKey(keyBytes, Convert.ToUInt32(srcHex, 16));
        // dst may be multi-codepoints; keep as UTF-16 string
        string s = HexToString(dstHex);
        if (_map.TryGetValue(key, out string? previous)) _retainedMappingTextCharacters -= previous.Length;
        else _retainedMappingTextCharacters += srcHex.Length;
        _retainedMappingTextCharacters += s.Length;
        _map[key] = s;
        uint code = (uint)key;
        if (keyBytes == 1) {
            (_oneByteMap ??= new string?[256])[code] = s;
        } else if (keyBytes == 2) {
            (_twoByteMap ??= new Dictionary<int, string>())[(int)code] = s;
        }
        if (!_reverseMapBudgetExhausted && s.Length > 0 && s.Length <= MaxReverseMappingTextCharacters) {
            ReverseMapNode node = _reverseMapRoot;
            for (int index = 0; index < s.Length; index++) {
                if (!node.TryGetOrAdd(s[index], ref _reverseMapNodeCount, MaxReverseMapNodes, out node)) {
                    _reverseMapBudgetExhausted = true;
                    return;
                }
            }

            if (node.CodeHex is null) {
                node.CodeHex = srcHex.ToUpperInvariant();
                // Reverse terminals and partially created branches survive forward-map replacement.
                // Charge their source strings separately, including conservative string-object overhead.
                _retainedReverseCodeCharacters += srcHex.Length;
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

    // ISO 32000-1, 9.10.3: a bfrange destination advances by incrementing its last byte. A multi-character
    // destination such as <005b00310035005d> ("[15]") is wider than 32 bits, so only its final UTF-16 unit advances.
    private static string OffsetRangeDestination(string dstHex, uint offset) {
        if (dstHex.Length <= 8) {
            uint first = dstHex.Length == 0 ? 0U : Convert.ToUInt32(dstHex, 16);
            return unchecked(first + offset).ToString("X", System.Globalization.CultureInfo.InvariantCulture);
        }

        if (offset == 0U) return dstHex;
        int lastUnit = dstHex.Length - 4;
        uint unit = unchecked(Convert.ToUInt32(dstHex.Substring(lastUnit), 16) + offset) & 0xFFFFU;
        return dstHex.Remove(lastUnit) + unit.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
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

    /// <summary>Maps exactly one code of <paramref name="length"/> bytes, without combining shorter entries.</summary>
    internal bool TryMapCode(byte[] bytes, int offset, int length, out string text) {
        text = string.Empty;
        if (length < 1 || length > MaxSourceCodeBytes || offset < 0 || offset + length > bytes.Length ||
            !TryGetMapping(bytes, offset, length, out string? mapped) || mapped is null) return false;
        text = mapped;
        return true;
    }

    /// <summary>Byte length of the longest source code.</summary>
    internal int MaxCodeBytes => _maxKeyBytes;

    /// <summary>Counts raw fallback bytes without materializing a second decoded text buffer.</summary>
    internal int CountUnmappedCodeBytes(byte[] bytes, Action? cancellationCheck) {
        int unmapped = 0;
        for (int index = 0; index < bytes.Length;) {
            cancellationCheck?.Invoke();
            int used = 0;
            for (int length = Math.Min(_maxKeyBytes, bytes.Length - index); length >= 1; length--) {
                if (TryGetMapping(bytes, index, length, out _)) {
                    used = length;
                    break;
                }
            }
            if (used == 0) { unmapped++; used = 1; }
            index += used;
        }
        return unmapped;
    }

    private string MapBytesCore(byte[] bytes, int maxOutputCharacters, bool allowUnmappedBytes, out bool allMapped) {
        if (maxOutputCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputCharacters), maxOutputCharacters, "Maximum decoded text characters must be positive.");
        }

        allMapped = true;
        System.Text.StringBuilder? sb = null;
        for (int i = 0; i < bytes.Length;) {
            // Greedy match up to _maxKeyBytes (1-2 bytes typical)
            int max = Math.Min(_maxKeyBytes, bytes.Length - i);
            string? mapped = null; int used = 0;
            for (int len = max; len >= 1; len--) {
                if (TryGetMapping(bytes, i, len, out var s)) { mapped = s; used = len; break; }
            }
            if (mapped is null) {
                allMapped = false;
                if (!allowUnmappedBytes) {
                    if (sb != null) ReleaseMapBuilder(sb);
                    return string.Empty;
                }
            }
            int appendLength = mapped?.Length ?? 1;
            long nextLength = (long)(sb?.Length ?? 0) + appendLength;
            if (nextLength > maxOutputCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, nextLength);
            }

            if (i == 0 && mapped is not null && used == bytes.Length) return mapped;
            sb ??= TakeMapBuilder();
            if (mapped is null) { sb.Append((char)bytes[i]); i++; } else { sb.Append(mapped); i += used; }
        }
        if (sb == null) return string.Empty;
        string result = sb.ToString();
        ReleaseMapBuilder(sb);
        return result;
    }

    // Every shown run is mapped through here, so one builder per thread is reused instead of growing a new one per run.
    [ThreadStatic]
    private static System.Text.StringBuilder? t_mapBuilder;
    private const int MaximumRetainedMapBuilderCapacity = 4096;

    private static System.Text.StringBuilder TakeMapBuilder() {
        System.Text.StringBuilder? builder = t_mapBuilder;
        if (builder == null) return new System.Text.StringBuilder();
        t_mapBuilder = null;
        builder.Clear();
        return builder;
    }

    private static void ReleaseMapBuilder(System.Text.StringBuilder builder) {
        if (builder.Capacity <= MaximumRetainedMapBuilderCapacity) t_mapBuilder = builder;
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

    private bool TryGetMapping(byte[] bytes, int start, int byteCount, out string? mapped) {
        if (byteCount == 1) {
            mapped = _oneByteMap?[bytes[start]];
            return mapped != null;
        }
        if (byteCount == 2) {
            mapped = null;
            return _twoByteMap != null && _twoByteMap.TryGetValue((bytes[start] << 8) | bytes[start + 1], out mapped);
        }
        return _map.TryGetValue(CreateKey(bytes, start, byteCount), out mapped);
    }

    private static readonly string[] KeyFormats = { "X0", "X2", "X4", "X6", "X8" };

    private static ulong CreateKey(int byteCount, uint code) => ((ulong)(uint)byteCount << 32) | code;

    private static ulong CreateKey(byte[] bytes, int start, int byteCount) {
        uint code = 0;
        for (int index = 0; index < byteCount; index++) {
            code = (code << 8) | bytes[start + index];
        }

        return CreateKey(byteCount, code);
    }

    private static string FormatKey(ulong key) =>
        ((uint)key).ToString(KeyFormats[(int)(key >> 32)], System.Globalization.CultureInfo.InvariantCulture);

}
