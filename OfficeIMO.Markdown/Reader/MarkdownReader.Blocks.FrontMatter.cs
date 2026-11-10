using System.Globalization;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Finds front matter at the top of the document: a <c>---</c> line and the next <c>---</c> line. Nothing between
    /// them is empty front matter, as Jekyll reads it. Other front matter formats, such as TOML between <c>+++</c>
    /// fences, would be recognized here.
    /// </summary>
    private static bool TryLocateFrontMatter(string[] lines, out int closingIndex) {
        closingIndex = -1;
        if (lines.Length == 0 || !IsFrontMatterFence(lines[0])) return false;
        for (int j = 1; j < lines.Length; j++) {
            if (IsFrontMatterFence(lines[j])) {
                closingIndex = j;
                return true;
            }
        }

        return false;
    }

    private static bool IsFrontMatterFence(string? line) => string.Equals(line?.Trim(), FrontMatterBlock.Fence, StringComparison.Ordinal);

    private static YamlLoadOptions CreateFrontMatterLoadOptions() {
        // Front matter comes from documents of unknown origin. Repeated keys are common in hand-edited files and the
        // last one wins in Jekyll, Hugo and the earlier OfficeIMO reader, so they are reported instead of failing.
        YamlLoadOptions options = YamlLoadOptions.UntrustedDefaults;
        options.DuplicateKeys = YamlDuplicateKeyPolicy.LastWins;
        return options;
    }

    /// <summary>Parses the lines between the fences with the YAML engine; invalid YAML is kept as text with a positioned error.</summary>
    private static FrontMatterBlock ParseFrontMatterBlock(string[] lines, int closingIndex, MarkdownReaderState state) {
        const int start = 1;
        int end = closingIndex - 1;
        string rawYaml = string.Join("\n", lines, start, end - start + 1);
        MarkdownSourceSpan? bodySpan = end < start ? null : CreateSpan(state, state.SourceLineOffset + start + 1, 1,
            state.SourceLineOffset + end + 1, Math.Max(1, lines[end].Length));
        int bodyLine = state.SourceLineOffset + start + 1;

        YamlDocument yaml;
        try {
            yaml = YamlDocument.Parse(rawYaml, CreateFrontMatterLoadOptions());
        } catch (YamlException exception) {
            int line = bodyLine + Math.Max(1, exception.Line) - 1;
            int column = Math.Max(1, exception.Column);
            string location = "line " + line.ToString(CultureInfo.InvariantCulture) + ", column " + column.ToString(CultureInfo.InvariantCulture);
            var diagnostic = new OfficeDiagnostic(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidYaml, OfficeDiagnosticSeverity.Error,
                "The front matter is not valid YAML and is kept as text (" + exception.Code + "): " + exception.Message,
                location, "Markdown");
            return FrontMatterBlock.FromInvalid(rawYaml, bodySpan, diagnostic, lines[0], lines[closingIndex]);
        }

        Dictionary<YamlNode, FrontMatterBlock.EntrySpans>? spans = state.CaptureSyntaxTree
            ? CreateFrontMatterEntrySpans(yaml, rawYaml, bodyLine, state)
            : null;
        return FrontMatterBlock.FromParsed(yaml, rawYaml, bodySpan, spans, lines[0], lines[closingIndex]);
    }

    private static Dictionary<YamlNode, FrontMatterBlock.EntrySpans>? CreateFrontMatterEntrySpans(YamlDocument yaml, string rawYaml, int bodyLine, MarkdownReaderState state) {
        if (yaml.Root is not YamlMapping mapping || mapping.Count == 0) return null;

        int[] lineStarts = GetLineStarts(rawYaml);
        var spans = new Dictionary<YamlNode, FrontMatterBlock.EntrySpans>(mapping.Count);
        for (int i = 0; i < mapping.Count; i++) {
            YamlMappingEntry entry = mapping[i];
            MarkdownSourceSpan? key = CreateYamlNodeSpan(entry.Key, rawYaml, lineStarts, bodyLine, state, out long keyStart, out _);
            MarkdownSourceSpan? value = CreateYamlNodeSpan(entry.Value, rawYaml, lineStarts, bodyLine, state, out _, out long valueEnd);
            MarkdownSourceSpan? whole = key.HasValue
                ? CreateYamlRangeSpan(keyStart, value.HasValue ? valueEnd : EndOf(entry.Key, rawYaml), rawYaml, lineStarts, bodyLine, state)
                : null;
            spans[entry.Key] = new FrontMatterBlock.EntrySpans(whole, key, value);
        }

        return spans;
    }

    private static long EndOf(YamlNode node, string rawYaml) => InclusiveEnd(node.Start.Offset, node.End.Offset, rawYaml);

    private static MarkdownSourceSpan? CreateYamlNodeSpan(YamlNode node, string rawYaml, int[] lineStarts, int bodyLine, MarkdownReaderState state,
        out long start, out long inclusiveEnd) {
        start = node.Start.Offset;
        inclusiveEnd = -1;
        if (node.Start.Line <= 0 || node.End.Offset <= start) return null;
        inclusiveEnd = InclusiveEnd(start, node.End.Offset, rawYaml);
        return inclusiveEnd < start ? null : CreateYamlRangeSpan(start, inclusiveEnd, rawYaml, lineStarts, bodyLine, state);
    }

    /// <summary>The last character of a node, before any line breaks its source extent ends with.</summary>
    private static long InclusiveEnd(long start, long exclusiveEnd, string rawYaml) {
        long end = Math.Min(exclusiveEnd, rawYaml.Length) - 1;
        while (end > start && (rawYaml[(int)end] == '\n' || rawYaml[(int)end] == '\r')) end--;
        return end;
    }

    private static MarkdownSourceSpan? CreateYamlRangeSpan(long start, long inclusiveEnd, string rawYaml, int[] lineStarts, int bodyLine, MarkdownReaderState state) {
        if (start < 0 || inclusiveEnd < start || inclusiveEnd >= rawYaml.Length) return null;
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
