namespace OfficeIMO.Email;

internal static class ContentLineCodec {
    internal static IReadOnlyList<ContentLineComponent> Parse(byte[] bytes, ContentLineReaderOptions options) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (bytes.LongLength > options.MaxInputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured byte limit.");
        return Parse(options.Encoding.GetString(bytes), options);
    }

    internal static IReadOnlyList<ContentLineComponent> Parse(string text, ContentLineReaderOptions options) {
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
                ContentLineProperty property = ParseProperty(line);
                if (property.Group == null && string.Equals(property.Name, "BEGIN", StringComparison.OrdinalIgnoreCase)) {
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

    internal static byte[] Serialize(IEnumerable<ContentLineComponent> components, ContentLineWriterOptions options) {
        if (components == null) throw new ArgumentNullException(nameof(components));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var output = new StringBuilder();
        long outputBytes = 0;
        foreach (ContentLineComponent component in components)
            WriteComponent(output, component, options, 1, ref outputBytes);
        byte[] bytes = options.Encoding.GetBytes(output.ToString());
        if (bytes.LongLength > options.MaxOutputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured output byte limit.");
        return bytes;
    }

    private static IEnumerable<string> Unfold(string text, ContentLineReaderOptions options) {
        StringBuilder? current = null;
        int currentBytes = 0;
        bool currentIsQuotedPrintable = false;
        bool firstPhysicalLine = true;
        int start = 0;
        for (int index = 0; index <= text.Length; index++) {
            if (index < text.Length && text[index] != '\r' && text[index] != '\n') continue;
            string physical = text.Substring(start, index - start);
            if (index < text.Length && text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            start = index + 1;
            if (firstPhysicalLine) {
                firstPhysicalLine = false;
                if (physical.Length > 0 && physical[0] == '\uFEFF') physical = physical.Substring(1);
            }
            if (current != null && current.Length > 0 && current[current.Length - 1] == '=' &&
                currentIsQuotedPrintable) {
                string continuation = physical.Length > 0 && (physical[0] == ' ' || physical[0] == '\t')
                    ? physical.Substring(1)
                    : physical;
                AppendUnfolded(current, "\r\n", ref currentBytes, options.MaxUnfoldedLineBytes);
                AppendUnfolded(current, continuation, ref currentBytes, options.MaxUnfoldedLineBytes);
            } else if (physical.Length > 0 && (physical[0] == ' ' || physical[0] == '\t') && current != null)
                AppendUnfolded(current, physical.Substring(1), ref currentBytes,
                    options.MaxUnfoldedLineBytes);
            else {
                if (current != null) yield return current.ToString();
                current = new StringBuilder(physical);
                currentBytes = Encoding.UTF8.GetByteCount(physical);
                if (currentBytes > options.MaxUnfoldedLineBytes)
                    throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");
                currentIsQuotedPrintable = IsQuotedPrintableContentLine(physical);
            }
        }
        if (current != null) yield return current.ToString();
    }

    private static void AppendUnfolded(StringBuilder current, string value, ref int currentBytes,
        int maximumBytes) {
        int addedBytes = Encoding.UTF8.GetByteCount(value);
        if (addedBytes > maximumBytes - currentBytes)
            throw new InvalidDataException("A content line exceeds the configured unfolded-line limit.");
        current.Append(value);
        currentBytes += addedBytes;
    }

    private static bool IsQuotedPrintableContentLine(string line) {
        int colon = FindDelimiter(line, ':');
        string header = colon >= 0 ? line.Substring(0, colon) : line;
        return header.IndexOf("ENCODING=QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) >= 0 ||
            header.IndexOf("ENCODING=QP", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static ContentLineProperty ParseProperty(string line) {
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
                parameter.Values.Add(DecodeParameter(Unquote(value), quoted));
            }
            property.Parameters.Add(parameter);
        }
        return property;
    }

    private static void WriteComponent(StringBuilder output, ContentLineComponent component,
        ContentLineWriterOptions options, int depth, ref long outputBytes) {
        if (component == null) throw new InvalidDataException("A null content-line component cannot be serialized.");
        if (depth > 256) throw new InvalidDataException("The content-line component graph is too deeply nested.");
        string name = ContentLineSyntax.RequireToken(component.Name, nameof(component.Name));
        AppendFolded(output, new[] { "BEGIN:", name }, options, ref outputBytes);
        foreach (ContentLineProperty property in component.Properties)
            AppendFolded(output, GetPropertyParts(property), options, ref outputBytes);
        foreach (ContentLineComponent child in component.Components)
            WriteComponent(output, child, options, depth + 1, ref outputBytes);
        AppendFolded(output, new[] { "END:", name }, options, ref outputBytes);
    }

    private static IEnumerable<string> GetPropertyParts(ContentLineProperty property) {
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
                yield return EncodeParameter(parameter.Values[index] ?? string.Empty);
            }
        }
        yield return ":";
        yield return property.Value ?? string.Empty;
    }

    private static void AppendFolded(StringBuilder output, IEnumerable<string> parts,
        ContentLineWriterOptions options, ref long outputBytes) {
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
                if (current.Length > 0 && octets + bytes > options.FoldAtOctets) {
                    EnsureOutputLimit(outputBytes + octets + newlineBytes, options);
                    output.Append(current).Append("\r\n");
                    outputBytes += octets + newlineBytes;
                    current.Clear();
                    current.Append(' ');
                    octets = options.Encoding.GetByteCount(" ");
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

    private static void EnsureOutputLimit(long projectedBytes, ContentLineWriterOptions options) {
        if (projectedBytes > options.MaxOutputBytes)
            throw new InvalidDataException("The content-line document exceeds the configured output byte limit.");
    }

    private static int FindDelimiter(string value, char delimiter) {
        bool quoted = false;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '"' && !IsEscaped(value, index)) quoted = !quoted;
            else if (value[index] == delimiter && !quoted) return index;
        }
        return -1;
    }

    private static IEnumerable<string> SplitDelimited(string value, char delimiter) {
        bool quoted = false;
        int start = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '"' && !IsEscaped(value, index)) quoted = !quoted;
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

    private static bool IsEscaped(string value, int index) {
        int backslashes = 0;
        for (int current = index - 1; current >= 0 && value[current] == '\\'; current--) backslashes++;
        return (backslashes & 1) != 0;
    }

    private static string DecodeParameter(string value, bool decodeLegacyQuotedBackslashes) {
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (decodeLegacyQuotedBackslashes && value[index] == '\\' && index + 1 < value.Length) {
                result.Append(value[++index]);
                continue;
            }
            if (value[index] != '^' || index + 1 >= value.Length) { result.Append(value[index]); continue; }
            char next = value[index + 1];
            if (next == '^') { result.Append('^'); index++; }
            else if (next == 'n' || next == 'N') { result.Append('\n'); index++; }
            else if (next == '\'') { result.Append('"'); index++; }
            else result.Append('^');
        }
        return result.ToString();
    }

    private static string EncodeParameter(string value) {
        string encoded = value.Replace("^", "^^").Replace("\r\n", "^n").Replace("\r", "^n")
            .Replace("\n", "^n").Replace("\"", "^'");
        bool quote = encoded.IndexOfAny(new[] { ':', ';', ',' }) >= 0 ||
            (encoded.Length > 0 && (char.IsWhiteSpace(encoded[0]) || char.IsWhiteSpace(encoded[encoded.Length - 1])));
        return quote ? "\"" + encoded + "\"" : encoded;
    }
}

internal static class ContentLineSyntax {
    internal static string RequireToken(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A content-line token cannot be empty.", parameterName);
        foreach (char character in value) {
            if (char.IsLetterOrDigit(character) || character == '-') continue;
            throw new ArgumentException("A content-line token contains an invalid character.", parameterName);
        }
        return value;
    }
}
