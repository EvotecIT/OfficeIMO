#nullable enable

namespace OfficeIMO.CSV;

internal sealed class CsvStreamingSource
{
    private readonly Func<TextReader> _readerFactory;
    private readonly CsvLoadOptions _options;
    private readonly int _skipRecordCount;

    public CsvStreamingSource(Func<TextReader> readerFactory, CsvLoadOptions options, int skipRecordCount)
    {
        _readerFactory = readerFactory;
        _options = options.Clone();
        _skipRecordCount = skipRecordCount;
    }

    public IEnumerable<object?[]> ReadRows()
    {
        using var reader = _readerFactory();
        var skipped = 0;
        foreach (var record in CsvParser.Parse(reader, _options))
        {
            if (skipped < _skipRecordCount)
            {
                skipped++;
                continue;
            }

            yield return record.Cast<object?>().ToArray();
        }
    }
}
