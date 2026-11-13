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
                OfficeDiagnostic tomlDiagnostic = exception.ToDiagnostic();
                var diagnostic = CreateFrontMatterDiagnostic(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidToml, "TOML",
                    tomlDiagnostic.Code, tomlDiagnostic.Message, exception.Line, exception.Column, bodyLine);
                return FrontMatterBlock.FromInvalid(MarkdownFrontMatterFormat.Toml, rawText, bodySpan, diagnostic, openingFence, closingFence);
            }

            Dictionary<string, FrontMatterBlock.EntrySpans>? keySpans = state.CaptureSyntaxTree
                ? CreateTomlEntrySpans(toml, rawText, bodyLine, state)
                : null;
            return FrontMatterBlock.FromParsed(toml, rawText, bodySpan, keySpans, openingFence, closingFence);
        }

        YamlDocument yaml;
        try {
            yaml = YamlDocument.Parse(rawText, CreateFrontMatterLoadOptions());
        } catch (YamlException exception) {
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
        // The engine diagnostic's message leaves out its position; the location is given in Markdown lines.
        int line = bodyLine + Math.Max(1, engineLine) - 1;
        int column = Math.Max(1, engineColumn);
        string location = "line " + line.ToString(CultureInfo.InvariantCulture) + ", column " + column.ToString(CultureInfo.InvariantCulture);
        return new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Error,
            "The front matter is not valid " + language + " and is kept as text (" + engineCode + "): " + message,
            location, "Markdown");
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
    /// Locates the top-level entries of TOML front matter from the model's key positions. A plain key/value pair spans its
    /// key and value; a table spans its header, and a table made by dotted keys or deeper headers spans the pair or header
    /// that first names it, without a value span.
    /// </summary>
    private static Dictionary<string, FrontMatterBlock.EntrySpans>? CreateTomlEntrySpans(TomlDocument toml, string rawText, int bodyLine, MarkdownReaderState state) {
        if (toml.Root.Count == 0) return null;

        int[] lineStarts = GetLineStarts(rawText);
        var spans = new Dictionary<string, FrontMatterBlock.EntrySpans>(toml.Root.Count, StringComparer.Ordinal);
        foreach (KeyValuePair<string, TomlNode> entry in toml.Root) {
            TomlNode node = entry.Value;
            TomlPosition keyStart = node.KeyStart;
            if (keyStart.Line <= 0) continue;
            MarkdownSourceSpan? key = CreateBodyRangeSpan(keyStart.Offset, node.KeyEnd.Offset - 1, rawText, lineStarts, bodyLine, state);
            TomlNode defining = FirstWithSource(node);
            if (defining.Start.Line <= 0) {
                spans[entry.Key] = new FrontMatterBlock.EntrySpans(key, key, value: null);
                continue;
            }

            long end = InclusiveEnd(defining.Start.Offset, defining.End.Offset, rawText);
            bool header = defining is TomlTable { TableKind: TomlTableKind.Header or TomlTableKind.ArrayItem };
            MarkdownSourceSpan? whole = CreateBodyRangeSpan(header ? defining.Start.Offset : keyStart.Offset, end, rawText, lineStarts, bodyLine, state);
            MarkdownSourceSpan? value = ReferenceEquals(defining, node) && !header
                ? CreateBodyRangeSpan(node.Start.Offset, end, rawText, lineStarts, bodyLine, state)
                : null;
            spans[entry.Key] = new FrontMatterBlock.EntrySpans(whole, key, value);
        }

        return spans;
    }

    /// <summary>The node, or for a table made by dotted keys or deeper headers its first entry with source text.</summary>
    private static TomlNode FirstWithSource(TomlNode node) {
        while (node.Start.Line <= 0) {
            TomlNode? next = node switch {
                TomlTable table => FirstEntry(table),
                TomlTableArray tables when tables.Count > 0 => tables[0],
                _ => null
            };
            if (next == null) break;
            node = next;
        }

        return node;
    }

    private static TomlNode? FirstEntry(TomlTable table) {
        foreach (KeyValuePair<string, TomlNode> entry in table) return entry.Value;
        return null;
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
