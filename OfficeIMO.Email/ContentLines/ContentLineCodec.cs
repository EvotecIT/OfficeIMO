namespace OfficeIMO.Email;

internal enum ContentLineParameterEncoding {
    Rfc6868,
    Legacy
}

internal static class ContentLineCodec {
    internal static IReadOnlyList<ContentLineComponent> Parse(byte[] bytes, ContentLineReaderOptions options,
        bool decodeRfc6868Parameters = true) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (bytes.LongLength > options.MaxInputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured byte limit.");
        return Parse(options.Encoding.GetString(bytes), options, decodeRfc6868Parameters);
    }

    internal static IReadOnlyList<ContentLineComponent> Parse(string text, ContentLineReaderOptions options,
        bool decodeRfc6868Parameters = true) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.Encoding.GetByteCount(text) > options.MaxInputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured byte limit.");

        try {
            var roots = new List<ContentLineComponent>();
            var stack = new Stack<ContentLineComponent>();
            int componentCount = 0;
            int propertyCount = 0;
            foreach (string line in Unfold(text, options)) {
                if (line.Length == 0) continue;
                ContentLineProperty property = ParseProperty(line, decodeRfc6868Parameters);
                if (property.Group == null && string.Equals(property.Name, "BEGIN", StringComparison.OrdinalIgnoreCase)) {
                    if (property.Parameters.Count != 0)
                        throw new InvalidDataException("BEGIN component delimiters cannot contain parameters.");
                    string name = ContentLineSyntax.RequireToken(property.Value, "componentName");
                    if (++componentCount > options.MaxComponents)
                        throw new InvalidDataException("The content-line document exceeds the configured component limit.");
                    if (stack.Count + 1 > options.MaxNestingDepth)
                        throw new InvalidDataException("The content-line document exceeds the configured nesting limit.");
                    var component = new ContentLineComponent(name);
                    if (stack.Count == 0) roots.Add(component); else stack.Peek().Components.Add(component);
                    stack.Push(component);
                    continue;
                }
                if (property.Group == null && string.Equals(property.Name, "END", StringComparison.OrdinalIgnoreCase)) {
                    if (property.Parameters.Count != 0)
                        throw new InvalidDataException("END component delimiters cannot contain parameters.");
                    if (stack.Count == 0 || !string.Equals(stack.Peek().Name, property.Value,
                        StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("The content-line document contains a mismatched END component.");
                    stack.Pop();
                    continue;
                }
                if (stack.Count == 0)
                    throw new InvalidDataException("A property was found outside a BEGIN/END component.");
                if (++propertyCount > options.MaxProperties)
                    throw new InvalidDataException("The content-line document exceeds the configured property limit.");
                stack.Peek().Properties.Add(property);
            }
            if (stack.Count != 0)
                throw new InvalidDataException("The content-line document contains an unterminated component.");
            return roots;
        } catch (ArgumentException exception) {
            throw new InvalidDataException("The content-line document contains an invalid token.", exception);
        }
    }

    internal static byte[] Serialize(IEnumerable<ContentLineComponent> components, ContentLineWriterOptions options,
        Func<ContentLineComponent, ContentLineParameterEncoding>? parameterEncoding = null) {
        if (components == null) throw new ArgumentNullException(nameof(components));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var output = new StringBuilder();
        long outputBytes = 0;
        var active = new HashSet<ContentLineComponent>();
        foreach (ContentLineComponent component in components) {
            ContentLineParameterEncoding encoding = parameterEncoding?.Invoke(component) ??
                ContentLineParameterEncoding.Rfc6868;
            WriteComponent(output, component, options, encoding, 1, active, ref outputBytes);
        }
        byte[] bytes = options.Encoding.GetBytes(output.ToString());
        if (bytes.LongLength > options.MaxOutputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured output byte limit.");
        return bytes;
    }

    private static IEnumerable<string> Unfold(string text, ContentLineReaderOptions options) {
        using (IEnumerator<string> physicalLines = EnumeratePhysicalLines(text).GetEnumerator()) {
            if (!physicalLines.MoveNext()) yield break;
            string physical = physicalLines.Current;
            while (true) {
                var current = new StringBuilder(physical);
                int currentBytes = Encoding.UTF8.GetByteCount(physical);
                if (currentBytes > options.MaxUnfoldedLineBytes && !EndsWithEquals(current))
                    throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");
                bool? currentIsQuotedPrintable = null;
                var deferredSoftBreaks = new List<int>();
                bool reachedEnd = true;
                while (physicalLines.MoveNext()) {
                    string next = physicalLines.Current;
                    bool folded = next.Length > 0 && (next[0] == ' ' || next[0] == '\t');
                    if (folded) {
                        if (EndsWithEquals(current)) {
                            if (currentIsQuotedPrintable == true) RemoveTrailingEquals(current, ref currentBytes);
                            else if (!currentIsQuotedPrintable.HasValue) {
                                int offset = current.Length - 1;
                                if (deferredSoftBreaks.Count == 0 ||
                                    deferredSoftBreaks[deferredSoftBreaks.Count - 1] != offset)
                                    deferredSoftBreaks.Add(offset);
                            }
                        }
                        AppendUnfolded(current, next.Substring(1), ref currentBytes,
                            options.MaxUnfoldedLineBytes,
                            currentIsQuotedPrintable.HasValue ? 0 : deferredSoftBreaks.Count);
                        continue;
                    }

                    ResolveQuotedPrintableState(current, deferredSoftBreaks,
                        ref currentIsQuotedPrintable, ref currentBytes);
                    if (currentIsQuotedPrintable == true && EndsWithEquals(current)) {
                        RemoveTrailingEquals(current, ref currentBytes);
                        AppendUnfolded(current, next, ref currentBytes, options.MaxUnfoldedLineBytes);
                        continue;
                    }
                    if (currentBytes > options.MaxUnfoldedLineBytes)
                        throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");

                    physical = next;
                    reachedEnd = false;
                    break;
                }

                ResolveQuotedPrintableState(current, deferredSoftBreaks,
                    ref currentIsQuotedPrintable, ref currentBytes);
                if (currentBytes > options.MaxUnfoldedLineBytes)
                    throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");
                yield return current.ToString();
                if (reachedEnd) yield break;
            }
        }
    }

    private static IEnumerable<string> EnumeratePhysicalLines(string text) {
        bool firstPhysicalLine = true;
        int start = 0;
        for (int index = 0; index <= text.Length; index++) {
            if (index < text.Length && text[index] != '\r' && text[index] != '\n') continue;
            string physical = text.Substring(start, index - start);
            if (index < text.Length && text[index] == '\r' && index + 1 < text.Length &&
                text[index + 1] == '\n') index++;
            start = index + 1;
            if (firstPhysicalLine) {
                firstPhysicalLine = false;
                if (physical.Length > 0 && physical[0] == '\uFEFF') physical = physical.Substring(1);
            }
            yield return physical;
        }
    }

    private static void AppendUnfolded(StringBuilder current, string value, ref int currentBytes,
        int maximumBytes, int deferredBytes = 0) {
        int addedBytes = Encoding.UTF8.GetByteCount(value);
        if (addedBytes > maximumBytes - (currentBytes - deferredBytes))
            throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");
        current.Append(value);
        currentBytes += addedBytes;
    }

    private static void ResolveQuotedPrintableState(StringBuilder current,
        List<int> deferredSoftBreaks, ref bool? currentIsQuotedPrintable, ref int currentBytes) {
        if (currentIsQuotedPrintable.HasValue) return;
        if (deferredSoftBreaks.Count == 0 && !EndsWithEquals(current)) {
            currentIsQuotedPrintable = false;
            return;
        }
        string line = current.ToString();
        int delimiter = FindDelimiter(line, ':');
        currentIsQuotedPrintable = delimiter > 0 &&
            IsQuotedPrintableHeader(line.Substring(0, delimiter));
        if (currentIsQuotedPrintable == true) {
            int writeIndex = 0;
            int deferredIndex = 0;
            for (int readIndex = 0; readIndex < current.Length; readIndex++) {
                if (deferredIndex < deferredSoftBreaks.Count &&
                    readIndex == deferredSoftBreaks[deferredIndex]) {
                    deferredIndex++;
                    if (readIndex > delimiter) {
                        currentBytes--;
                        continue;
                    }
                }
                if (writeIndex != readIndex) current[writeIndex] = current[readIndex];
                writeIndex++;
            }
            current.Length = writeIndex;
        }
        deferredSoftBreaks.Clear();
    }

    private static bool EndsWithEquals(StringBuilder value) =>
        value.Length > 0 && value[value.Length - 1] == '=';

    private static void RemoveTrailingEquals(StringBuilder value, ref int currentBytes) {
        value.Length--;
        currentBytes--;
    }

    private static bool IsQuotedPrintableHeader(string header) {
        try {
            foreach (string segment in SplitDelimited(header, ';').Skip(1)) {
                int equals = FindDelimiter(segment, '=');
                if (equals <= 0 || !string.Equals(segment.Substring(0, equals), "ENCODING",
                    StringComparison.OrdinalIgnoreCase)) continue;
                foreach (string value in SplitDelimited(segment.Substring(equals + 1), ',')) {
                    string encoding = Unquote(value);
                    if (string.Equals(encoding, "QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(encoding, "QP", StringComparison.OrdinalIgnoreCase)) return true;
                }
            }
        } catch (InvalidDataException) {
            // Malformed quoting is rejected later by the full property parser.
        }
        return false;
    }

    private static ContentLineProperty ParseProperty(string line, bool decodeRfc6868Parameters) {
        int colon = FindDelimiter(line, ':');
        if (colon <= 0) throw new InvalidDataException("A content line does not contain a property/value delimiter.");
        string[] segments = SplitDelimited(line.Substring(0, colon), ';').ToArray();
        string propertyToken = segments[0];
        string? group = null;
        int dot = propertyToken.IndexOf('.');
        if (dot >= 0) {
            group = ContentLineSyntax.RequireToken(propertyToken.Substring(0, dot), "group");
            propertyToken = propertyToken.Substring(dot + 1);
        }
        var property = new ContentLineProperty(propertyToken, line.Substring(colon + 1)) { Group = group };
        for (int index = 1; index < segments.Length; index++) {
            string segment = segments[index];
            int equals = FindDelimiter(segment, '=');
            if (equals < 0) {
                property.Parameters.Add(new ContentLineParameter(segment));
                continue;
            }
            if (equals == 0) throw new InvalidDataException("A content-line parameter is malformed.");
            var parameter = new ContentLineParameter(segment.Substring(0, equals));
            foreach (string value in SplitDelimited(segment.Substring(equals + 1), ',')) {
                bool quoted = value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"';
                parameter.Values.Add(DecodeParameter(Unquote(value), quoted, decodeRfc6868Parameters));
            }
            property.Parameters.Add(parameter);
        }
        return property;
    }

    private static void WriteComponent(StringBuilder output, ContentLineComponent component,
        ContentLineWriterOptions options, ContentLineParameterEncoding parameterEncoding,
        int depth, ISet<ContentLineComponent> active, ref long outputBytes) {
        if (component == null) throw new InvalidDataException("A null content-line component cannot be serialized.");
        if (depth > ContentLineComponent.MaximumTraversalDepth)
            throw new InvalidDataException("The content-line component graph is too deeply nested.");
        if (!active.Add(component))
            throw new InvalidDataException("The content-line component graph contains a cycle.");
        try {
            string name = ContentLineSyntax.RequireToken(component.Name, nameof(component.Name));
            AppendFolded(output, new[] { "BEGIN:", name }, options, ref outputBytes);
            foreach (ContentLineProperty property in component.Properties)
                AppendFolded(output, GetPropertyParts(property, parameterEncoding), options, ref outputBytes,
                    avoidQuotedPrintableSoftBreak: IsQuotedPrintableProperty(property));
            foreach (ContentLineComponent child in component.Components)
                WriteComponent(output, child, options, parameterEncoding, depth + 1, active, ref outputBytes);
            AppendFolded(output, new[] { "END:", name }, options, ref outputBytes);
        } finally {
            active.Remove(component);
        }
    }

    private static IEnumerable<string> GetPropertyParts(ContentLineProperty property,
        ContentLineParameterEncoding parameterEncoding) {
        if (property == null) throw new InvalidDataException("A null content-line property cannot be serialized.");
        if (!string.IsNullOrEmpty(property.Group)) {
            yield return ContentLineSyntax.RequireToken(property.Group!, nameof(property.Group));
            yield return ".";
        }
        yield return ContentLineSyntax.RequireToken(property.Name, nameof(property.Name));
        foreach (ContentLineParameter parameter in property.Parameters) {
            yield return ";";
            yield return ContentLineSyntax.RequireToken(parameter.Name, nameof(parameter.Name));
            if (parameter.Values.Count == 0) continue;
            yield return "=";
            for (int index = 0; index < parameter.Values.Count; index++) {
                if (index > 0) yield return ",";
                yield return EncodeParameter(parameter.Values[index] ?? string.Empty, parameterEncoding);
            }
        }
        yield return ":";
        string value = property.Value ?? string.Empty;
        foreach (char character in value) {
            if (character < '\u0020' && character != '\t' || character == '\u007F') {
                throw new InvalidDataException(
                    "A content-line property value cannot contain an ASCII control character other than HTAB; " +
                    "use the format escape for textual line breaks.");
            }
        }
        yield return value;
    }

    private static void AppendFolded(StringBuilder output, IEnumerable<string> parts,
        ContentLineWriterOptions options, ref long outputBytes,
        bool avoidQuotedPrintableSoftBreak = false) {
        var current = new StringBuilder();
        int octets = 0;
        int newlineBytes = options.Encoding.GetByteCount("\r\n");
        foreach (string part in parts) {
            if (part == null) continue;
            for (int index = 0; index < part.Length;) {
                int length = char.IsHighSurrogate(part[index]) && index + 1 < part.Length &&
                    char.IsLowSurrogate(part[index + 1]) ? 2 : 1;
                string character = part.Substring(index, length);
                int bytes = options.Encoding.GetByteCount(character);
                if (current.Length == 0 && bytes > options.FoldAtOctets)
                    throw new InvalidDataException(
                        "The configured fold limit cannot accommodate an encoded character.");
                if (current.Length > 0 && octets + bytes > options.FoldAtOctets) {
                    bool moveEqualsToContinuation = avoidQuotedPrintableSoftBreak &&
                        current[current.Length - 1] == '=';
                    int equalsBytes = 0;
                    if (moveEqualsToContinuation) {
                        equalsBytes = options.Encoding.GetByteCount("=");
                    }
                    int continuationBytes = options.Encoding.GetByteCount(" ") + equalsBytes;
                    if (continuationBytes + bytes > options.FoldAtOctets)
                        throw new InvalidDataException(
                            "The configured fold limit cannot accommodate a continuation prefix and encoded character.");
                    if (moveEqualsToContinuation) {
                        current.Length--;
                        octets -= equalsBytes;
                    }
                    EnsureOutputLimit(outputBytes + octets + newlineBytes, options);
                    output.Append(current).Append("\r\n");
                    outputBytes += octets + newlineBytes;
                    current.Clear();
                    current.Append(' ');
                    octets = continuationBytes - equalsBytes;
                    if (moveEqualsToContinuation) {
                        current.Append('=');
                        octets += equalsBytes;
                    }
                }
                EnsureOutputLimit(outputBytes + octets + bytes + newlineBytes, options);
                current.Append(character);
                octets += bytes;
                index += length;
            }
        }
        EnsureOutputLimit(outputBytes + octets + newlineBytes, options);
        output.Append(current).Append("\r\n");
        outputBytes += octets + newlineBytes;
    }

    private static bool IsQuotedPrintableProperty(ContentLineProperty property) =>
        property.Parameters.Where(parameter => string.Equals(
                parameter.Name, "ENCODING", StringComparison.OrdinalIgnoreCase))
            .SelectMany(parameter => parameter.Values)
            .Any(value => string.Equals(value, "QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(value, "QP", StringComparison.OrdinalIgnoreCase));

    private static void EnsureOutputLimit(long projectedBytes, ContentLineWriterOptions options) {
        if (projectedBytes > options.MaxOutputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured output byte limit.");
    }

    private static int FindDelimiter(string value, char delimiter) {
        bool quoted = false;
        bool hasAmbiguousQuotes = HasNonStandardEscapedQuoteCandidate(value);
        bool[]? canFinishUnquoted = null;
        bool[]? canFinishQuoted = null;
        if (hasAmbiguousQuotes) {
            BuildQuoteReachability(value, delimiter, out canFinishUnquoted, out canFinishQuoted);
        }
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '"') {
                if (hasAmbiguousQuotes && IsBackslashQuote(value, index) &&
                    ShouldTreatAsEscapedQuote(quoted, index,
                        canFinishUnquoted!, canFinishQuoted!)) continue;
                quoted = !quoted;
            }
            else if (value[index] == delimiter && !quoted) return index;
        }
        return -1;
    }

    private static IEnumerable<string> SplitDelimited(string value, char delimiter) {
        bool quoted = false;
        int start = 0;
        bool hasAmbiguousQuotes = HasNonStandardEscapedQuoteCandidate(value);
        bool[]? canFinishUnquoted = null;
        bool[]? canFinishQuoted = null;
        if (hasAmbiguousQuotes) {
            BuildQuoteReachability(value, null, out canFinishUnquoted, out canFinishQuoted);
        }
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '"') {
                if (hasAmbiguousQuotes && IsBackslashQuote(value, index) &&
                    ShouldTreatAsEscapedQuote(quoted, index,
                        canFinishUnquoted!, canFinishQuoted!)) continue;
                quoted = !quoted;
            }
            else if (value[index] == delimiter && !quoted) {
                yield return value.Substring(start, index - start);
                start = index + 1;
            }
        }
        if (quoted) throw new InvalidDataException("A quoted content-line parameter is unterminated.");
        yield return value.Substring(start);
    }

    private static string Unquote(string value) => value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"'
        ? value.Substring(1, value.Length - 2)
        : value;

    private static bool HasNonStandardEscapedQuoteCandidate(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '"' && IsBackslashQuote(value, index)) return true;
        }
        return false;
    }

    private static bool IsBackslashQuote(string value, int index) =>
        index > 0 && value[index - 1] == '\\';

    private static void BuildQuoteReachability(string value, char? terminalDelimiter,
        out bool[] canFinishUnquoted, out bool[] canFinishQuoted) {
        canFinishUnquoted = new bool[value.Length + 1];
        canFinishQuoted = new bool[value.Length + 1];
        if (!terminalDelimiter.HasValue) canFinishUnquoted[value.Length] = true;
        for (int index = value.Length - 1; index >= 0; index--) {
            if (terminalDelimiter.HasValue && value[index] == terminalDelimiter.Value) {
                canFinishUnquoted[index] = true;
                canFinishQuoted[index] = canFinishQuoted[index + 1];
            } else if (value[index] == '"') {
                if (IsBackslashQuote(value, index)) {
                    bool reachable = canFinishUnquoted[index + 1] || canFinishQuoted[index + 1];
                    canFinishUnquoted[index] = reachable;
                    canFinishQuoted[index] = reachable;
                } else {
                    canFinishUnquoted[index] = canFinishQuoted[index + 1];
                    canFinishQuoted[index] = canFinishUnquoted[index + 1];
                }
            } else {
                canFinishUnquoted[index] = canFinishUnquoted[index + 1];
                canFinishQuoted[index] = canFinishQuoted[index + 1];
            }
        }
    }

    private static bool ShouldTreatAsEscapedQuote(bool quoted, int index,
        bool[] canFinishUnquoted, bool[] canFinishQuoted) {
        int next = index + 1;
        bool canStay = quoted ? canFinishQuoted[next] : canFinishUnquoted[next];
        bool canToggle = quoted ? canFinishUnquoted[next] : canFinishQuoted[next];
        return quoted ? canStay || !canToggle : !canToggle && canStay;
    }

    private static string DecodeParameter(string value, bool decodeLegacyQuotedBackslashes,
        bool decodeRfc6868Parameters) {
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (decodeLegacyQuotedBackslashes && value[index] == '\\' && index + 1 < value.Length) {
                char escaped = value[index + 1];
                if (escaped == '"') {
                    result.Append(escaped);
                    index++;
                    continue;
                }
            }
            if (!decodeRfc6868Parameters || value[index] != '^' || index + 1 >= value.Length) {
                result.Append(value[index]);
                continue;
            }
            char next = value[index + 1];
            if (next == '^') { result.Append('^'); index++; }
            else if (next == 'n' || next == 'N') { result.Append('\n'); index++; }
            else if (next == '\'') { result.Append('"'); index++; }
            else result.Append('^');
        }
        return result.ToString();
    }

    internal static void DecodeRfc6868Parameters(ContentLineComponent component) {
        foreach (ContentLineProperty property in component.Properties) {
            foreach (ContentLineParameter parameter in property.Parameters) {
                for (int index = 0; index < parameter.Values.Count; index++) {
                    parameter.Values[index] = DecodeParameter(parameter.Values[index] ?? string.Empty,
                        decodeLegacyQuotedBackslashes: false, decodeRfc6868Parameters: true);
                }
            }
        }
        foreach (ContentLineComponent child in component.Components)
            DecodeRfc6868Parameters(child);
    }

    private static string EncodeParameter(string value, ContentLineParameterEncoding parameterEncoding) {
        if (parameterEncoding == ContentLineParameterEncoding.Legacy) {
            if (value.IndexOfAny(new[] { '\r', '\n' }) >= 0) {
                throw new InvalidDataException(
                    "A legacy content-line parameter contains a line break that its syntax cannot represent.");
            }
            bool legacyQuote = value.IndexOfAny(new[] { ':', ';', ',', '"' }) >= 0 ||
                (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1])));
            string legacyEncoded = value.Replace("\"", "\\\"");
            return legacyQuote ? "\"" + legacyEncoded + "\"" : legacyEncoded;
        }
        string encoded = value.Replace("^", "^^").Replace("\r\n", "^n").Replace("\r", "^n")
            .Replace("\n", "^n").Replace("\"", "^'");
        bool quote = encoded.IndexOfAny(new[] { ':', ';', ',' }) >= 0 ||
            (encoded.Length > 0 && (char.IsWhiteSpace(encoded[0]) || char.IsWhiteSpace(encoded[encoded.Length - 1])));
        return quote ? "\"" + encoded + "\"" : encoded;
    }
}

internal static class ContentLineSyntax {
    internal static bool IsToken(string? value) {
        if (value == null || value.Length == 0) return false;
        foreach (char character in value) {
            if (character >= 'A' && character <= 'Z' ||
                character >= 'a' && character <= 'z' ||
                character >= '0' && character <= '9' ||
                character == '-') continue;
            return false;
        }
        return true;
    }

    internal static string RequireToken(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("A content-line token cannot be empty.", parameterName);
        if (!IsToken(value))
            throw new ArgumentException("A content-line token contains an invalid character.", parameterName);
        return value;
    }
}
