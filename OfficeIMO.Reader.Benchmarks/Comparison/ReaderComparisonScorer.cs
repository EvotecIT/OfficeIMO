namespace OfficeIMO.Reader.Benchmarks.Comparison;

internal static class ReaderComparisonScorer {
    public static IReadOnlyList<ReaderComparisonProbeResult> ScoreMarkdown(
        string markdown,
        IReadOnlyList<ReaderComparisonProbe> probes,
        bool rejected) => probes.Select(probe => ScoreMarkdownProbe(markdown, probe, rejected)).ToArray();

    public static IReadOnlyList<ReaderComparisonProbeResult> ScoreOfficeDocument(
        string markdown,
        OfficeDocumentReadResult? document,
        IReadOnlyList<ReaderComparisonProbe> probes,
        bool rejected) => probes.Select(probe => ScoreOfficeProbe(markdown, document, probe, rejected)).ToArray();

    private static ReaderComparisonProbeResult ScoreMarkdownProbe(
        string markdown,
        ReaderComparisonProbe probe,
        bool rejected) {
        bool applied = IsMarkdownProbe(probe.Kind) || probe.Kind == ReaderComparisonProbeKind.RejectsMalformedInput;
        bool passed = applied && EvaluateMarkdown(markdown, probe, rejected);
        return Result(probe, applied, passed);
    }

    private static ReaderComparisonProbeResult ScoreOfficeProbe(
        string markdown,
        OfficeDocumentReadResult? document,
        ReaderComparisonProbe probe,
        bool rejected) {
        if (IsMarkdownProbe(probe.Kind) || probe.Kind == ReaderComparisonProbeKind.RejectsMalformedInput) {
            return Result(probe, true, EvaluateMarkdown(markdown, probe, rejected));
        }

        bool passed = document != null && probe.Kind switch {
            ReaderComparisonProbeKind.RichTable => document.Tables.Count > 0 || document.Chunks.Any(HasTable),
            ReaderComparisonProbeKind.RichLink => document.Links.Count > 0,
            ReaderComparisonProbeKind.RichAsset => document.Assets.Count > 0,
            ReaderComparisonProbeKind.LocationPath => document.Chunks.Any(chunk => !string.IsNullOrWhiteSpace(chunk.Location.Path)),
            ReaderComparisonProbeKind.LocationHeading => document.Chunks.Any(chunk => !string.IsNullOrWhiteSpace(chunk.Location.HeadingPath)),
            ReaderComparisonProbeKind.LocationSheet => document.Chunks.Any(chunk => !string.IsNullOrWhiteSpace(chunk.Location.Sheet)),
            ReaderComparisonProbeKind.LocationSlide => document.Chunks.Any(chunk => chunk.Location.Slide.HasValue),
            ReaderComparisonProbeKind.LocationPage => document.Chunks.Any(chunk => chunk.Location.Page.HasValue),
            _ => false
        };
        return Result(probe, true, passed);
    }

    private static bool EvaluateMarkdown(string markdown, ReaderComparisonProbe probe, bool rejected) {
        if (probe.Kind == ReaderComparisonProbeKind.RejectsMalformedInput) return rejected;
        if (string.IsNullOrWhiteSpace(markdown)) return false;

        return probe.Kind switch {
            ReaderComparisonProbeKind.ContainsText => Contains(markdown, probe.Marker),
            ReaderComparisonProbeKind.MarkdownHeading => Lines(markdown).Any(line =>
                line.TrimStart().StartsWith("#", StringComparison.Ordinal) && Contains(line, probe.Marker)),
            ReaderComparisonProbeKind.MarkdownListItem => Lines(markdown).Any(line =>
                IsListItem(line.TrimStart()) && Contains(line, probe.Marker)),
            ReaderComparisonProbeKind.MarkdownTable => Lines(markdown).Any(line =>
                line.Contains('|') && Contains(line, probe.Marker)),
            ReaderComparisonProbeKind.MarkdownLink => Contains(markdown, "[" + probe.Marker + "]("),
            ReaderComparisonProbeKind.MarkdownImage => Contains(markdown, "![" + probe.Marker + "]("),
            _ => false
        };
    }

    private static bool IsMarkdownProbe(ReaderComparisonProbeKind kind) => kind is
        ReaderComparisonProbeKind.ContainsText or
        ReaderComparisonProbeKind.MarkdownHeading or
        ReaderComparisonProbeKind.MarkdownListItem or
        ReaderComparisonProbeKind.MarkdownTable or
        ReaderComparisonProbeKind.MarkdownLink or
        ReaderComparisonProbeKind.MarkdownImage;

    private static bool HasTable(ReaderChunk chunk) => chunk.Tables != null && chunk.Tables.Count > 0;

    private static bool IsListItem(string line) {
        if (line.StartsWith("- ", StringComparison.Ordinal) ||
            line.StartsWith("* ", StringComparison.Ordinal) ||
            line.StartsWith("+ ", StringComparison.Ordinal)) return true;
        int dot = line.IndexOf('.', StringComparison.Ordinal);
        return dot > 0 && dot < 10 && line.Substring(0, dot).All(char.IsDigit) &&
            dot + 1 < line.Length && char.IsWhiteSpace(line[dot + 1]);
    }

    private static IEnumerable<string> Lines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');

    private static bool Contains(string value, string marker) =>
        value.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0;

    private static ReaderComparisonProbeResult Result(
        ReaderComparisonProbe probe,
        bool applied,
        bool passed) => new ReaderComparisonProbeResult {
            Id = probe.Id,
            Kind = probe.Kind.ToString(),
            Applied = applied,
            Passed = passed
        };
}