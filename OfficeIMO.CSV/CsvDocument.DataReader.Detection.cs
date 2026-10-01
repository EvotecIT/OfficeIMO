#nullable enable
#if NET8_0_OR_GREATER
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private const int IncrementalDelimiterDetectionCharacterLimit = 64 * 1024;

    private static async Task<TextReader> DetectIncrementalDelimiterAsync(TextReader reader,
        CsvLoadOptions options, CancellationToken token)
    {
        var text = new StringBuilder();
        var buffer = new char[4096];
        int physicalLines = 0, nextProbe = DelimiterDetectionSampleLimit;
        bool pendingSample = false, pendingInQuotes = false, previousCr = false, ended = false;
        while (text.Length < IncrementalDelimiterDetectionCharacterLimit)
        {
            token.ThrowIfCancellationRequested();
            int count = await reader.ReadAsync(buffer.AsMemory(0,
                Math.Min(buffer.Length, IncrementalDelimiterDetectionCharacterLimit - text.Length)), token).ConfigureAwait(false);
            if (count == 0) { ended = true; break; }
            text.Append(buffer, 0, count);
            for (int i = 0; i < count; i++)
            {
                char current = buffer[i];
                if (pendingSample && current == '"') pendingInQuotes = !pendingInQuotes;
                if (current == '\r' || current == '\n' && !previousCr) physicalLines++;
                previousCr = current == '\r';
            }
            if (physicalLines < nextProbe || pendingSample && pendingInQuotes) continue;

            string lookahead = text.ToString();
            int completeLength = GetCompleteDetectionPrefixLength(lookahead);
            string[] candidates = ReadIncrementalDelimiterSamples(lookahead, completeLength, options);
            int completeSamples = candidates.Count(IsLogicalDelimiterDetectionRecordComplete);
            if (completeSamples >= DelimiterDetectionSampleLimit) break;
            // Canonical sampling determines comment recovery and logical-record completeness.
            // While its final sample is quoted and incomplete, only track quote parity until
            // it can complete; this avoids repeatedly reparsing a long multiline prefix.
            pendingSample = candidates.Length > 0 && !IsLogicalDelimiterDetectionRecordComplete(candidates[candidates.Length - 1]);
            pendingInQuotes = pendingSample;
            if (pendingSample)
                UpdateLogicalDelimiterDetectionQuoteState(lookahead.Substring(completeLength), ref pendingInQuotes);
            nextProbe = physicalLines + Math.Max(1, DelimiterDetectionSampleLimit - completeSamples);
        }

        string prefix = text.ToString();
        // A capped partial physical line is replayed, but never treated as a complete sample.
        int sampleLength = ended ? prefix.Length : GetCompleteDetectionPrefixLength(prefix);
        string[] samples = ReadIncrementalDelimiterSamples(prefix, sampleLength, options)
            .Where(IsLogicalDelimiterDetectionRecordComplete).ToArray();
        options.Delimiter = SelectDetectedDelimiter(samples, options);
        options.DetectDelimiter = false;
        token.ThrowIfCancellationRequested();
        return new CsvPrefixTextReader(prefix, reader);
    }

    private static int GetCompleteDetectionPrefixLength(string prefix) =>
        Math.Max(prefix.LastIndexOf('\n'), prefix.LastIndexOf('\r')) + 1;

    private static string[] ReadIncrementalDelimiterSamples(string prefix, int length, CsvLoadOptions options)
    {
        using var reader = new StringReader(prefix.Substring(0, length));
        return ReadDelimiterDetectionSamples(reader, options, useHeaderDiscovery: true).ToArray();
    }
}
#endif
