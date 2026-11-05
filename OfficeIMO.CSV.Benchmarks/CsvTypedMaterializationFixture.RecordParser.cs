using System.Globalization;
using RecordParser.Builders.Reader;
using RecordParser.Extensions;
using RecordParser.Parsers;

namespace OfficeIMO.CSV.Benchmarks;

internal sealed partial class CsvTypedMaterializationFixture {
    private static readonly IVariableLengthReader<CsvBenchmarkRow> RecordParserReader =
        new VariableLengthReaderBuilder<CsvBenchmarkRow>()
            .Map(row => row.Id, 0)
            .Map(row => row.Name, 1)
            .Map(row => row.Department, 2)
            .Map(row => row.Region, 3)
            .Map(row => row.IsEnabled, 4)
            .Map(row => row.Created, 5, text => DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind))
            .Map(row => row.Score, 6)
            .Map(row => row.Owner, 7)
            .Map(row => row.TicketCount, 8)
            .Map(row => row.Notes, 9)
            .Build(",", CultureInfo.InvariantCulture);

    internal CsvBenchmarkRow[] RecordParserSequential() => ReadRecordParser(parallel: false, degreeOfParallelism: 1);

    internal CsvBenchmarkRow[] RecordParserParallel(int degreeOfParallelism) => ReadRecordParser(parallel: true, degreeOfParallelism);

    internal void ValidateRecordParser(CsvBenchmarkRow[] rows) => Validate(nameof(RecordParserParallel), rows);

    private CsvBenchmarkRow[] ReadRecordParser(bool parallel, int degreeOfParallelism) {
        using var reader = new StringReader(_csvText);
        // The compiled ordinal map is reusable, just as the other lanes' typed mappings are.
        // Verify that this operation's source still has the layout that map requires.
        string? header = reader.ReadLine();
        if (header != "Id,Name,Department,Region,IsEnabled,Created,Score,Owner,TicketCount,Notes") {
            throw new InvalidDataException("RecordParser received an unexpected header layout.");
        }
        return reader.ReadRecords(RecordParserReader, new VariableLengthReaderOptions {
            HasHeader = false,
            ContainsQuotedFields = true,
            ParallelismOptions = new() {
                Enabled = parallel,
                EnsureOriginalOrdering = true,
                MaxDegreeOfParallelism = degreeOfParallelism
            }
        }).ToArray();
    }
}
