#nullable enable

using System.Data;
using System.Globalization;

namespace OfficeIMO.CSV.Benchmarks;

internal static class DataTableBenchmarkUtilities
{
    public static int Measure(DataTable table)
    {
        var checksum = 0;
        foreach (DataRow row in table.Rows)
        {
            foreach (DataColumn column in table.Columns)
            {
                var value = row[column];
                checksum += 1 + (value == DBNull.Value
                    ? 0
                    : Convert.ToString(value, CultureInfo.InvariantCulture)?.Length ?? 0);
            }
        }

        return checksum;
    }
}
