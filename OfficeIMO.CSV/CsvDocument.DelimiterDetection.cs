#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private const int DelimiterDetectionSampleLimit = 64;

    private static readonly char[] DefaultDelimiterCandidates = { ',', ';', '|', '\t' };

    private static CsvLoadOptions ResolveLoadOptions(Func<TextReader> readerFactory, CsvLoadOptions options, bool useHeaderDiscoveryForDelimiterDetection = true)
    {
        if (!options.DetectDelimiter)
        {
            return options;
        }

        var resolved = options.Clone();
        resolved.Delimiter = DetectDelimiter(readerFactory, options, useHeaderDiscoveryForDelimiterDetection);
        resolved.DetectDelimiter = false;
        return resolved;
    }

    private static char DetectDelimiter(Func<TextReader> readerFactory, CsvLoadOptions options, bool useHeaderDiscovery)
    {
        var candidates = options.DelimiterCandidates is { Length: > 0 }
            ? options.DelimiterCandidates
            : DefaultDelimiterCandidates;

        using var reader = readerFactory();
        var samples = ReadDelimiterDetectionSamples(reader, options, useHeaderDiscovery).ToArray();
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

    private static IEnumerable<string> ReadDelimiterDetectionSamples(TextReader reader, CsvLoadOptions options, bool useHeaderDiscovery)
    {
        var candidates = options.DelimiterCandidates is { Length: > 0 }
            ? options.DelimiterCandidates
            : DefaultDelimiterCandidates;
        var recordsToSkip = GetInitialRecordsToSkip(options);
        var allowPreHeaderCommentSkip = true;
        using var records = ReadLogicalDelimiterDetectionRecords(
            reader,
            line => ShouldSkipCommentDuringDelimiterDetection(line, options, useHeaderDiscovery, allowPreHeaderCommentSkip),
            line => IsDelimiterDetectionHeaderCandidate(line, options, candidates)).GetEnumerator();
        while (records.MoveNext())
        {
            var record = records.Current;
            if (IsBlankDelimiterDetectionRecord(record, options))
            {
                if (options.AllowEmptyLines)
                {
                    if (recordsToSkip > 0)
                    {
                        recordsToSkip--;
                        continue;
                    }

                    yield return record;
                    allowPreHeaderCommentSkip = false;
                    break;
                }

                continue;
            }

            if (ShouldSkipCommentDuringDelimiterDetection(record, options, useHeaderDiscovery, allowPreHeaderCommentSkip))
            {
                continue;
            }

            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                continue;
            }

            yield return record;
            allowPreHeaderCommentSkip = false;
            break;
        }

        var count = 1;
        while (count < DelimiterDetectionSampleLimit && records.MoveNext())
        {
            var record = records.Current;
            if (IsBlankDelimiterDetectionRecord(record, options) && !options.AllowEmptyLines)
            {
                continue;
            }

            if (ShouldSkipCommentDuringDelimiterDetection(record, options, useHeaderDiscovery, allowPreHeaderCommentSkip))
            {
                continue;
            }

            yield return record;
            count++;
        }
    }

    private static bool IsBlankDelimiterDetectionRecord(string record, CsvLoadOptions options) =>
        record.Length == 0 || (options.TrimWhitespace && record.Trim().Length == 0);

    private static IEnumerable<string> ReadLogicalDelimiterDetectionRecords(
        TextReader reader,
        Func<string, bool> shouldSkipRawCommentRecord,
        Func<string, bool> isHeaderCandidate)
    {
        var pendingLines = new Queue<string>();
        while (TryReadDelimiterDetectionLine(reader, pendingLines, out var line))
        {
            var inQuotes = false;
            UpdateLogicalDelimiterDetectionQuoteState(line, ref inQuotes);
            if (shouldSkipRawCommentRecord(line) && inQuotes)
            {
                SkipRawDelimiterDetectionCommentRecord(reader, pendingLines, line, isHeaderCandidate);
                continue;
            }

            if (!inQuotes)
            {
                yield return line;
                continue;
            }

            var record = new StringBuilder(line);
            while (TryReadDelimiterDetectionLine(reader, pendingLines, out line))
            {
                record.Append('\n');
                record.Append(line);
                UpdateLogicalDelimiterDetectionQuoteState(line, ref inQuotes);
                if (!inQuotes)
                {
                    break;
                }
            }

            yield return record.ToString();
        }
    }

    private static bool TryReadDelimiterDetectionLine(TextReader reader, Queue<string> pendingLines, out string line)
    {
        if (pendingLines.Count > 0)
        {
            line = pendingLines.Dequeue();
            return true;
        }

        var next = reader.ReadLine();
        if (next is null)
        {
            line = string.Empty;
            return false;
        }

        line = next;
        return true;
    }

    private static void SkipRawDelimiterDetectionCommentRecord(
        TextReader reader,
        Queue<string> pendingLines,
        string firstLine,
        Func<string, bool> isHeaderCandidate)
    {
        var continuations = new List<string>();
        var inQuotes = false;
        UpdateLogicalDelimiterDetectionQuoteState(firstLine, ref inQuotes);
        while (reader.ReadLine() is { } next)
        {
            continuations.Add(next);
            UpdateLogicalDelimiterDetectionQuoteState(next, ref inQuotes);
            if (!inQuotes)
            {
                return;
            }

            if (!isHeaderCandidate(firstLine) && isHeaderCandidate(next))
            {
                EnqueueDelimiterDetectionContinuations(pendingLines, continuations);
                return;
            }
        }

        EnqueueDelimiterDetectionContinuations(pendingLines, continuations);
    }

    private static void EnqueueDelimiterDetectionContinuations(Queue<string> pendingLines, List<string> continuations)
    {
        foreach (var continuation in continuations)
        {
            pendingLines.Enqueue(continuation);
        }
    }

    private static bool IsLogicalDelimiterDetectionRecordComplete(string record)
    {
        var inQuotes = false;
        UpdateLogicalDelimiterDetectionQuoteState(record, ref inQuotes);
        return !inQuotes;
    }

    private static void UpdateLogicalDelimiterDetectionQuoteState(string text, ref bool inQuotes)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '"')
            {
                continue;
            }

            if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
            {
                i++;
                continue;
            }

            inQuotes = !inQuotes;
        }
    }

    private static bool ShouldSkipCommentDuringDelimiterDetection(string line, CsvLoadOptions options, bool useHeaderDiscovery, bool allowPreHeaderCommentSkip)
    {
        if (line.Length == 0 || line[0] != options.CommentCharacter)
        {
            return false;
        }

        var canReadW3CFieldsHeader = useHeaderDiscovery &&
            allowPreHeaderCommentSkip &&
            options.HasHeaderRow &&
            options.Header is null &&
            options.RecognizeW3CFieldsHeader;

        var skipPreHeaderComment = allowPreHeaderCommentSkip &&
            useHeaderDiscovery &&
            options.HasHeaderRow &&
            options.Header is null &&
            options.SkipCommentRowsBeforeHeader;

        if (!options.SkipCommentRows && !skipPreHeaderComment)
        {
            return false;
        }

        return !canReadW3CFieldsHeader || !IsW3CFieldsLine(line, options);
    }

    private static bool IsW3CFieldsLine(string line, CsvLoadOptions options) =>
        options.RecognizeW3CFieldsHeader && line.StartsWith("#Fields:", StringComparison.OrdinalIgnoreCase);

    private static bool IsDelimiterDetectionHeaderCandidate(string line, CsvLoadOptions options, IReadOnlyList<char> candidates) =>
        IsW3CFieldsLine(line, options) || candidates.Any(line.Contains);

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
