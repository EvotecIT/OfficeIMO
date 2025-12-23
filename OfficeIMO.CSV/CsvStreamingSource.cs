#nullable enable

namespace OfficeIMO.CSV;

internal sealed class CsvStreamingSource
{
    private readonly Func<TextReader> _readerFactory;
    private readonly CsvLoadOptions _options;
    private readonly bool _skipFirstRecord;

    public CsvStreamingSource(Func<TextReader> readerFactory, CsvLoadOptions options, bool skipFirstRecord)
    {
        _readerFactory = readerFactory;
        _options = options.Clone();
        _skipFirstRecord = skipFirstRecord;
    }

    public IEnumerable<object?[]> ReadRows()
    {
        using var reader = _readerFactory();
        var first = true;
        foreach (var record in CsvParser.Parse(reader, _options))
        {
            if (_skipFirstRecord && first)
            {
                first = false;
                continue;
            }

            first = false;
            yield return record.Cast<object?>().ToArray();
        }
    }
}
