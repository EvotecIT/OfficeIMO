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

    public static int Measure(IDataReader reader)
    {
        var checksum = 0;
        while (reader.Read())
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var value = reader.GetValue(i);
                checksum += 1 + (value == DBNull.Value
                    ? 0
                    : Convert.ToString(value, CultureInfo.InvariantCulture)?.Length ?? 0);
            }
        }

        return checksum;
    }
}
