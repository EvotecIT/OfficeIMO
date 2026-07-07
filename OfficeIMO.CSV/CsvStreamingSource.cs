#nullable enable

namespace OfficeIMO.CSV;

internal sealed class CsvStreamingSource
{
    private readonly Func<TextReader> _readerFactory;
    private readonly CsvLoadOptions _options;
    private readonly int _skipRecordCount;
    private readonly int _headerCount;

    public CsvStreamingSource(Func<TextReader> readerFactory, CsvLoadOptions options, int skipRecordCount, int headerCount)
    {
        _readerFactory = readerFactory;
        _options = options.Clone();
        _skipRecordCount = skipRecordCount;
        _headerCount = headerCount;
    }

    public CsvLoadOptions Options => _options;

    public int SourceColumnCount => _headerCount - (_options.StaticColumns?.Count ?? 0);

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

            yield return CsvDocument.BuildParsedObjectValues(record, _headerCount, _options);
        }
    }

    public IEnumerable<object?[]> ReadReusableRows()
    {
        using var reader = _readerFactory();
        var skipped = 0;
        object?[]? row = null;
        foreach (var record in CsvParser.ParseReusable(reader, _options))
        {
            if (skipped < _skipRecordCount)
            {
                skipped++;
                continue;
            }

            row = CsvDocument.FillParsedObjectValues(record, _headerCount, _options, row);
            yield return row;
        }
    }

    public IEnumerable<IReadOnlyList<string>> ReadReusableStringRows()
    {
        using var reader = _readerFactory();
        var skipped = 0;
        foreach (var record in CsvParser.ParseReusable(reader, _options))
        {
            if (skipped < _skipRecordCount)
            {
                skipped++;
                continue;
            }

            yield return record;
        }
    }
}
