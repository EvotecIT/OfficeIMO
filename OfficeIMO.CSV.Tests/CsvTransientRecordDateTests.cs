#if NET8_0_OR_GREATER
using System;
using System.Linq;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvTransientRecordDateTests {
    [Theory]
    [InlineData(1)]
    [InlineData(4)]
    public void DefaultDateFormatsParseThroughSequentialAndWorkerBatches(int degree) {
        const string text = "Date,Name\n2024-01-02,Ada\n2024-01-03,Grace\n";
        DateTime[] dates = CsvDocument.ReadTextRowsAsParallel<DateTime>(text, header => {
            int date = header.GetOrdinal("Date");
            return row => row.GetDateTime(date);
        }, parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = degree }).ToArray();

        Assert.Equal(new[] { new DateTime(2024, 1, 2), new DateTime(2024, 1, 3) }, dates);
    }
}
#endif
