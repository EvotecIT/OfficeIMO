using System.Globalization;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Finds front matter at the top of the document: a <c>---</c> line and the next <c>---</c> line (YAML), or a
    /// <c>+++</c> line and the next <c>+++</c> line (TOML, as Hugo reads it). Nothing between them is empty front matter,
    /// as Jekyll reads it.
    /// </summary>
    internal static bool TryLocateFrontMatter(string[] lines, out int closingIndex) {
        closingIndex = -1;
        if (lines.Length == 0) return false;
        string? fence = FrontMatterFence(lines[0]);
        if (fence == null) return false;
        for (int j = 1; j < lines.Length; j++) {
            if (string.Equals(FrontMatterFence(lines[j]), fence, StringComparison.Ordinal)) {
                closingIndex = j;
                return true;
            }
        }

        return false;
    }

    // Fences start at column 0, as in Jekyll and Hugo; an indented "---" is indented code or a thematic break.
    private static string? FrontMatterFence(string? line) {
        string? trimmed = line?.TrimEnd();
        if (string.Equals(trimmed, FrontMatterBlock.YamlFence, StringComparison.Ordinal)) return FrontMatterBlock.YamlFence;
        if (string.Equals(trimmed, FrontMatterBlock.TomlFence, StringComparison.Ordinal)) return FrontMatterBlock.TomlFence;
        return null;
    }

    // Front matter comes from documents of unknown origin. Repeated YAML keys are common in hand-edited files; the
    // engine's default lets the last one win, as Jekyll and Hugo do, and reports each repeat instead of failing.
    private static YamlLoadOptions CreateFrontMatterLoadOptions() => YamlLoadOptions.UntrustedDefaults;

    /// <summary>
    /// Parses the front matter of a Markdown source whose line endings are already <c>\n</c>, without parsing the rest of
    /// the document. Returns <see langword="null"/> when the source does not start with front matter.
    /// </summary>
    /// <param name="markdown">The Markdown source.</param>
    /// <param name="closingIndex">The 0-based line of the closing fence, or -1.</param>
    internal static FrontMatterBlock? ParseFrontMatter(string markdown, out int closingIndex) {
        string[] lines = markdown.Split('\n');
        if (!TryLocateFrontMatter(lines, out closingIndex)) return null;
        return ParseFrontMatterBlock(lines, closingIndex, new MarkdownReaderState());
    }

    /// <summary>
    /// Parses the lines between the fences with the YAML or TOML engine; invalid input is kept as text with a positioned
    /// error.
    /// </summary>
    private static FrontMatterBlock ParseFrontMatterBlock(string[] lines, int closingIndex, MarkdownReaderState state) {
        const int start = 1;
        int end = closingIndex - 1;
        string rawText = string.Join("\n", lines, start, end - start + 1);
        MarkdownSourceSpan? bodySpan = end < start ? null : CreateSpan(state, state.SourceLineOffset + start + 1, 1,
            state.SourceLineOffset + end + 1, Math.Max(1, lines[end].Length));
        int bodyLine = state.SourceLineOffset + start + 1;
        string openingFence = lines[0];
        string closingFence = lines[closingIndex];

        if (FrontMatterFence(openingFence) == FrontMatterBlock.TomlFence) {
            TomlDocument toml;
            try {
                toml = TomlDocument.Parse(rawText, TomlLoadOptions.UntrustedDefaults);
            } catch (TomlException exception) {
                var diagnostic = CreateFrontMatterDiagnostic(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidToml, "TOML",
                    exception.Code, WithoutPosition(exception), exception.Line, exception.Column, bodyLine);
                return FrontMatterBlock.FromInvalid(MarkdownFrontMatterFormat.Toml, rawText, bodySpan, diagnostic, openingFence, closingFence);
            }

            Dictionary<string, FrontMatterBlock.EntrySpans>? keySpans = state.CaptureSyntaxTree
                ? CreateTomlEntrySpans(rawText, bodyLine, state)
                : null;
            return FrontMatterBlock.FromParsed(toml, rawText, bodySpan, keySpans, openingFence, closingFence);
        }

        YamlDocument yaml;
        try {
            yaml = YamlDocument.Parse(rawText, CreateFrontMatterLoadOptions());
        } catch (YamlException exception) {
            // The YAML diagnostic's message leaves out its YAML-relative position; the location is given in Markdown lines.
            OfficeDiagnostic yamlDiagnostic = exception.ToDiagnostic();
            var diagnostic = CreateFrontMatterDiagnostic(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidYaml, "YAML",
                yamlDiagnostic.Code, yamlDiagnostic.Message, exception.Line, exception.Column, bodyLine);
            return FrontMatterBlock.FromInvalid(MarkdownFrontMatterFormat.Yaml, rawText, bodySpan, diagnostic, openingFence, closingFence);
        }

        Dictionary<YamlNode, FrontMatterBlock.EntrySpans>? spans = state.CaptureSyntaxTree
            ? CreateFrontMatterEntrySpans(yaml, rawText, bodyLine, state)
            : null;
        return FrontMatterBlock.FromParsed(yaml, rawText, bodySpan, spans, openingFence, closingFence);
    }

    private static OfficeDiagnostic CreateFrontMatterDiagnostic(string code, string language, string engineCode, string message,
        int engineLine, int engineColumn, int bodyLine) {
        int line = bodyLine + Math.Max(1, engineLine) - 1;
        int column = Math.Max(1, engineColumn);
        string location = "line " + line.ToString(CultureInfo.InvariantCulture) + ", column " + column.ToString(CultureInfo.InvariantCulture);
        return new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Error,
            "The front matter is not valid " + language + " and is kept as text (" + engineCode + "): " + message,
            location, "Markdown");
    }

    /// <summary>The exception's message without the TOML-relative position it ends with; the diagnostic gives Markdown lines.</summary>
    private static string WithoutPosition(TomlException exception) {
        string message = exception.Message;
        if (exception.Line <= 0) return message;
        string suffix = " (line " + exception.Line.ToString(CultureInfo.InvariantCulture) + ", column "
            + exception.Column.ToString(CultureInfo.InvariantCulture) + ")";
        return message.EndsWith(suffix, StringComparison.Ordinal) ? message.Substring(0, message.Length - suffix.Length) : message;
    }

    private static Dictionary<YamlNode, FrontMatterBlock.EntrySpans>? CreateFrontMatterEntrySpans(YamlDocument yaml, string rawText, int bodyLine, MarkdownReaderState state) {
        if (yaml.Root is not YamlMapping mapping || mapping.Count == 0) return null;

        int[] lineStarts = GetLineStarts(rawText);
        var spans = new Dictionary<YamlNode, FrontMatterBlock.EntrySpans>(mapping.Count);
        for (int i = 0; i < mapping.Count; i++) {
            YamlMappingEntry entry = mapping[i];
            MarkdownSourceSpan? key = CreateYamlNodeSpan(entry.Key, rawText, lineStarts, bodyLine, state, out long keyStart, out _);
            MarkdownSourceSpan? value = CreateYamlNodeSpan(entry.Value, rawText, lineStarts, bodyLine, state, out _, out long valueEnd);
            MarkdownSourceSpan? whole = key.HasValue
                ? CreateBodyRangeSpan(keyStart, value.HasValue ? valueEnd : EndOf(entry.Key, rawText), rawText, lineStarts, bodyLine, state)
                : null;
            spans[entry.Key] = new FrontMatterBlock.EntrySpans(whole, key, value);
        }

        return spans;
    }

    /// <summary>
    /// Locates the top-level entries of TOML front matter: the key/value pairs before the first header, and the first
    /// header of each top-level table. A dotted key or a table gets the span of its first key segment and of the pair or
    /// header that first names it, without a value span; a plain pair also gets its value's span.
    /// </summary>
    private static Dictionary<string, FrontMatterBlock.EntrySpans>? CreateTomlEntrySpans(string rawText, int bodyLine, MarkdownReaderState state) {
        if (rawText.Length == 0) return null;

        int[] lineStarts = GetLineStarts(rawText);
        var spans = new Dictionary<string, FrontMatterBlock.EntrySpans>(StringComparer.Ordinal);
        var reader = new TomlReader(rawText);
        bool inHeader = false, afterHeader = false, inValue = false;
        string? key = null;
        long headerStart = -1, keyStart = -1, keyEnd = -1, valueStart = -1;
        int segments = 0;

        while (reader.Read()) {
            switch (reader.TokenType) {
                case TomlTokenType.TableHeaderStart:
                case TomlTokenType.ArrayOfTablesHeaderStart:
                    inHeader = true;
                    afterHeader = true;
                    headerStart = reader.Start.Offset;
                    key = null;
                    break;
                case TomlTokenType.TableHeaderEnd:
                case TomlTokenType.ArrayOfTablesHeaderEnd:
                    inHeader = false;
                    if (key != null && !spans.ContainsKey(key)) {
                        spans[key] = new FrontMatterBlock.EntrySpans(
                            CreateBodyRangeSpan(headerStart, reader.End.Offset - 1, rawText, lineStarts, bodyLine, state),
                            CreateBodyRangeSpan(keyStart, keyEnd - 1, rawText, lineStarts, bodyLine, state),
                            value: null);
                    }

                    key = null;
                    break;
                case TomlTokenType.Key when inHeader:
                    if (key == null) {
                        key = reader.GetString();
                        keyStart = reader.Start.Offset;
                        keyEnd = reader.End.Offset;
                    }

                    break;
                case TomlTokenType.Key when !afterHeader && reader.Depth == 0:
                    if (key == null) {
                        key = reader.GetString();
                        keyStart = reader.Start.Offset;
                        keyEnd = reader.End.Offset;
                        segments = 0;
                    }

                    segments++;
                    break;
                case TomlTokenType.Comment:
                case TomlTokenType.Key:
                    break;
                default:
                    if (key == null || afterHeader || reader.Depth != 0) break;
                    if (!inValue) valueStart = reader.Start.Offset;
                    if (reader.TokenType is TomlTokenType.StartArray or TomlTokenType.StartInlineTable) {
                        inValue = true;
                        break;
                    }

                    inValue = false;
                    long valueEnd = InclusiveEnd(valueStart, reader.End.Offset, rawText);
                    if (!spans.ContainsKey(key)) {
                        spans[key] = new FrontMatterBlock.EntrySpans(
                            CreateBodyRangeSpan(keyStart, valueEnd, rawText, lineStarts, bodyLine, state),
                            CreateBodyRangeSpan(keyStart, keyEnd - 1, rawText, lineStarts, bodyLine, state),
                            segments == 1 ? CreateBodyRangeSpan(valueStart, valueEnd, rawText, lineStarts, bodyLine, state) : null);
                    }

                    key = null;
                    break;
            }
        }

        return spans;
    }

    private static long EndOf(YamlNode node, string rawText) => InclusiveEnd(node.Start.Offset, node.End.Offset, rawText);

    private static MarkdownSourceSpan? CreateYamlNodeSpan(YamlNode node, string rawText, int[] lineStarts, int bodyLine, MarkdownReaderState state,
        out long start, out long inclusiveEnd) {
        start = node.Start.Offset;
        inclusiveEnd = -1;
        if (node.Start.Line <= 0 || node.End.Offset <= start) return null;
        inclusiveEnd = InclusiveEnd(start, node.End.Offset, rawText);
        return inclusiveEnd < start ? null : CreateBodyRangeSpan(start, inclusiveEnd, rawText, lineStarts, bodyLine, state);
    }

    /// <summary>The last character of a node, before any line breaks its source extent ends with.</summary>
    private static long InclusiveEnd(long start, long exclusiveEnd, string rawText) {
        long end = Math.Min(exclusiveEnd, rawText.Length) - 1;
        while (end > start && (rawText[(int)end] == '\n' || rawText[(int)end] == '\r')) end--;
        return end;
    }

    private static MarkdownSourceSpan? CreateBodyRangeSpan(long start, long inclusiveEnd, string rawText, int[] lineStarts, int bodyLine, MarkdownReaderState state) {
        if (start < 0 || inclusiveEnd < start || inclusiveEnd >= rawText.Length) return null;
        int startLine = FindLine(lineStarts, (int)start);
        int endLine = FindLine(lineStarts, (int)inclusiveEnd);
        return CreateSpan(state,
            bodyLine + startLine, (int)start - lineStarts[startLine] + 1,
            bodyLine + endLine, (int)inclusiveEnd - lineStarts[endLine] + 1);
    }

    private static int[] GetLineStarts(string text) {
        var starts = new List<int> { 0 };
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\n') starts.Add(i + 1);
        }

        return starts.ToArray();
    }

    private static int FindLine(int[] lineStarts, int offset) {
        int index = Array.BinarySearch(lineStarts, offset);
        return index >= 0 ? index : ~index - 1;
    }
}
