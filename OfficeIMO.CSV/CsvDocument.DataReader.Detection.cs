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
        int records = 0;
        bool inQuotes = false, previousCr = false, ended = false;
        while (text.Length < IncrementalDelimiterDetectionCharacterLimit && records < DelimiterDetectionSampleLimit)
        {
            token.ThrowIfCancellationRequested();
            int count = await reader.ReadAsync(buffer.AsMemory(0,
                Math.Min(buffer.Length, IncrementalDelimiterDetectionCharacterLimit - text.Length)), token).ConfigureAwait(false);
            if (count == 0) { ended = true; break; }
            text.Append(buffer, 0, count);
            for (int i = 0; i < count; i++)
            {
                char current = buffer[i];
                if (current == '"') inQuotes = !inQuotes;
                if (!inQuotes && (current == '\r' || current == '\n' && !previousCr)) records++;
                previousCr = current == '\r';
            }
        }

        string prefix = text.ToString();
        // A capped partial physical line is replayed, but never treated as a complete sample.
        int sampleLength = ended ? prefix.Length : Math.Max(prefix.LastIndexOf('\n'), prefix.LastIndexOf('\r')) + 1;
        using var sampleReader = new StringReader(prefix.Substring(0, sampleLength));
        string[] samples = ReadDelimiterDetectionSamples(sampleReader, options, useHeaderDiscovery: true)
            .Where(IsLogicalDelimiterDetectionRecordComplete).ToArray();
        options.Delimiter = SelectDetectedDelimiter(samples, options);
        options.DetectDelimiter = false;
        token.ThrowIfCancellationRequested();
        return new CsvPrefixTextReader(prefix, reader);
    }
}
#endif
