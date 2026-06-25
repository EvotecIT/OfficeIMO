#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private const int DelimiterDetectionSampleLimit = 64;

    private static readonly char[] DefaultDelimiterCandidates = { ',', ';', '|', '\t' };

    private static CsvLoadOptions ResolveLoadOptions(Func<TextReader> readerFactory, CsvLoadOptions options)
    {
        if (!options.DetectDelimiter)
        {
            return options;
        }

        var resolved = options.Clone();
        resolved.Delimiter = DetectDelimiter(readerFactory, options);
        resolved.DetectDelimiter = false;
        return resolved;
    }

    private static char DetectDelimiter(Func<TextReader> readerFactory, CsvLoadOptions options)
    {
        var candidates = options.DelimiterCandidates is { Length: > 0 }
            ? options.DelimiterCandidates
            : DefaultDelimiterCandidates;

        using var reader = readerFactory();
        var samples = ReadDelimiterDetectionSamples(reader, options).ToArray();
        if (samples.Length == 0)
        {
            return options.Delimiter;
        }

        var bestDelimiter = options.Delimiter;
        var bestScore = DelimiterScore.Empty;

        foreach (var candidate in candidates)
        {
            var score = ScoreDelimiter(samples, candidate);
            if (score.CompareTo(bestScore) > 0)
            {
                bestDelimiter = candidate;
                bestScore = score;
            }
        }

        return bestScore.FirstLineFieldCount > 1 ? bestDelimiter : options.Delimiter;
    }

    private static IEnumerable<string> ReadDelimiterDetectionSamples(TextReader reader, CsvLoadOptions options)
    {
        var recordsToSkip = GetInitialRecordsToSkip(options);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (line.Length == 0)
            {
                if (options.AllowEmptyLines)
                {
                    yield return line;
                    break;
                }

                continue;
            }

            if (ShouldSkipCommentDuringDelimiterDetection(line, options))
            {
                continue;
            }

            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                continue;
            }

            yield return line;
            break;
        }

        var count = 1;
        while (count < DelimiterDetectionSampleLimit && (line = reader.ReadLine()) is not null)
        {
            if (line.Length == 0 && !options.AllowEmptyLines)
            {
                continue;
            }

            yield return line;
            count++;
        }
    }

    private static bool ShouldSkipCommentDuringDelimiterDetection(string line, CsvLoadOptions options) =>
        (options.SkipCommentRows || (options.HasHeaderRow && options.Header is null && options.SkipCommentRowsBeforeHeader)) &&
        line.Length > 0 &&
        line[0] == options.CommentCharacter &&
        !IsW3CFieldsLine(line, options);

    private static bool IsW3CFieldsLine(string line, CsvLoadOptions options) =>
        options.RecognizeW3CFieldsHeader && line.StartsWith("#Fields:", StringComparison.OrdinalIgnoreCase);

    private static DelimiterScore ScoreDelimiter(IReadOnlyList<string> samples, char delimiter)
    {
        var firstCount = CountFields(samples[0], delimiter);
        var matchingLines = 0;
        var usableLines = 0;
        var totalFields = 0;
        var minFields = int.MaxValue;
        var maxFields = 0;

        foreach (var sample in samples)
        {
            var count = CountFields(sample, delimiter);
            if (count <= 1)
            {
                continue;
            }

            usableLines++;
            totalFields += count;
            minFields = Math.Min(minFields, count);
            maxFields = Math.Max(maxFields, count);

            if (count == firstCount)
            {
                matchingLines++;
            }
        }

        var spread = usableLines == 0 ? int.MaxValue : maxFields - minFields;
        return new DelimiterScore(firstCount, matchingLines, usableLines, totalFields, spread);
    }

    private static int CountFields(string line, char delimiter)
    {
        var count = 1;
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var current = line[i];
            if (current == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    i++;
                    continue;
                }

                inQuotes = !inQuotes;
                continue;
            }

            if (!inQuotes && current == delimiter)
            {
                count++;
            }
        }

        return count;
    }

    private readonly struct DelimiterScore : IComparable<DelimiterScore>
    {
        public static readonly DelimiterScore Empty = new(1, 0, 0, 0, int.MaxValue);

        public DelimiterScore(int firstLineFieldCount, int matchingLines, int usableLines, int totalFields, int spread)
        {
            FirstLineFieldCount = firstLineFieldCount;
            MatchingLines = matchingLines;
            UsableLines = usableLines;
            TotalFields = totalFields;
            Spread = spread;
        }

        public int FirstLineFieldCount { get; }

        private int MatchingLines { get; }

        private int UsableLines { get; }

        private int TotalFields { get; }

        private int Spread { get; }

        public int CompareTo(DelimiterScore other)
        {
            var comparison = FirstLineFieldCount.CompareTo(other.FirstLineFieldCount);
            if (comparison != 0)
            {
                return comparison;
            }

            comparison = MatchingLines.CompareTo(other.MatchingLines);
            if (comparison != 0)
            {
                return comparison;
            }

            comparison = UsableLines.CompareTo(other.UsableLines);
            if (comparison != 0)
            {
                return comparison;
            }

            comparison = other.Spread.CompareTo(Spread);
            if (comparison != 0)
            {
                return comparison;
            }

            return TotalFields.CompareTo(other.TotalFields);
        }
    }
}
