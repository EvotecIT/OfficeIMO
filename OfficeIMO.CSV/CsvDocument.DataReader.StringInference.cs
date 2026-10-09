#nullable enable
using System.Threading;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private CsvDataReader CreateStreamingStringInferredDataReader(
        int sampleSize, CsvLoadOptions options, CancellationToken cancellationToken)
    {
        var rows = _streamingSource!.ReadReusableStringRows(options).GetEnumerator();
        try
        {
            var sampled = new List<IReadOnlyList<string>>(Math.Min(sampleSize, 4096));
            var inferred = new InferredColumn[_header.Count];
            for (int ordinal = 0; ordinal < inferred.Length; ordinal++)
                inferred[ordinal] = new InferredColumn(_header[ordinal]);
            object?[]? values = null;
            while (sampled.Count < sampleSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                options.OperationCancellationToken = cancellationToken;
                bool hasRow;
                try { hasRow = rows.MoveNext(); }
                finally { options.OperationCancellationToken = default; }
                if (!hasRow) break;
                cancellationToken.ThrowIfCancellationRequested();
                // Keep the original text and row width for the eventual reader. Inference
                // observes normalized values without replacing the borrowed-record input.
                sampled.Add(rows.Current.ToArray());
                values = FillParsedObjectValues(rows.Current, _header.Count, options, values);
                for (int ordinal = 0; ordinal < inferred.Length; ordinal++)
                    inferred[ordinal].Observe(values[ordinal], _culture, _dateTimeFormats);
            }
            cancellationToken.ThrowIfCancellationRequested();
            var schema = new CsvSchema(inferred.Select(column => column.ToSchemaColumn(sampled.Count)).ToArray());
            var owner = new CsvStreamingStringDataReaderRowOwner(rows);
            return new CsvDataReader(CreateDataReaderColumns(_header, schema),
                EnumerateSampledThenRemainingStringRows(sampled, owner),
                _streamingSource.SourceColumnCount, options, _culture, _dateTimeFormats, owner);
        }
        catch
        {
            rows.Dispose();
            throw;
        }
    }

    private static IEnumerable<IReadOnlyList<string>> EnumerateSampledThenRemainingStringRows(
        IReadOnlyList<IReadOnlyList<string>> sampled,
        CsvStreamingStringDataReaderRowOwner remaining)
    {
        try
        {
            foreach (var row in sampled) yield return row;
            while (remaining.MoveNext()) yield return remaining.Current;
        }
        finally { remaining.Dispose(); }
    }

    private sealed class CsvStreamingStringDataReaderRowOwner : IDisposable
    {
        private IEnumerator<IReadOnlyList<string>>? _rows;
        internal CsvStreamingStringDataReaderRowOwner(IEnumerator<IReadOnlyList<string>> rows) => _rows = rows;
        internal IReadOnlyList<string> Current => _rows!.Current;
        internal bool MoveNext() => _rows?.MoveNext() == true;
        public void Dispose()
        {
            _rows?.Dispose();
            _rows = null;
        }
    }
}
