using System.Data.Common;
using OfficeIMO.Tabular;
using Sylvan.Data;
using SylvanCsvReader = Sylvan.Data.Csv.CsvDataReader;
using SylvanExcelReader = Sylvan.Data.Excel.ExcelDataReader;

namespace OfficeIMO.Tabular.Benchmarks;

internal static class TabularBenchmarkOperations {
    internal static Observation ReadSylvanCsvStrings() {
        using var text = File.OpenText(FixtureData.CsvPath);
        using var reader = SylvanCsvReader.Create(text);
        return ReadStrings(reader);
    }

    internal static Observation ReadOfficeCsvStrings() {
        using var reader = TabularReader.Open(FixtureData.CsvPath);
        return ReadStrings(reader);
    }

    internal static Observation ReadSylvanCsvTyped() {
        using var text = File.OpenText(FixtureData.CsvPath);
        using var reader = SylvanCsvReader.Create(text);
        return ReadTyped(reader);
    }

    internal static Observation ReadOfficeCsvTyped() {
        using var reader = TabularReader.Open(FixtureData.CsvPath);
        return ReadTyped(reader);
    }

    internal static Observation ReadSylvanXlsxTyped() {
        using var reader = SylvanExcelReader.Create(FixtureData.XlsxPath);
        return ReadTyped(reader);
    }

    internal static Observation ReadOfficeXlsxTyped() {
        using var reader = TabularReader.Open(FixtureData.XlsxPath);
        return ReadTyped(reader);
    }

    internal static Observation ReadSylvanXlsbTyped() {
        using var reader = SylvanExcelReader.Create(FixtureData.XlsbPath);
        return ReadTyped(reader);
    }

    internal static Observation ReadOfficeXlsbTyped() {
        using var reader = TabularReader.Open(FixtureData.XlsbPath);
        return ReadTyped(reader);
    }

    internal static Observation ReadSylvanXlsxRecords() {
        using var reader = SylvanExcelReader.Create(FixtureData.XlsxPath);
        return ObserveRecords(reader.GetRecords<SalesRecord>());
    }

    internal static Observation ReadOfficeXlsxRecords() {
        using var reader = TabularReader.Open(FixtureData.XlsxPath);
        return ObserveRecords(reader.ReadRecords<SalesRecord>());
    }

    private static Observation ReadStrings(DbDataReader reader) {
        int rows = 0;
        int cells = 0;
        long checksum = 17;
        int fieldCount = reader.FieldCount;
        while (reader.Read()) {
            rows++;
            for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                string value = reader.GetString(ordinal);
                cells++;
                checksum = unchecked((checksum * 31) + StringChecksum(value));
            }
        }

        return new Observation(rows, cells, checksum);
    }

    private static Observation ReadTyped(DbDataReader reader) {
        int rows = 0;
        int cells = 0;
        long checksum = 17;
        while (reader.Read()) {
            var record = new SalesRecord {
                Region = reader.GetString(0),
                Country = reader.GetString(1),
                ItemType = reader.GetString(2),
                SalesChannel = reader.GetString(3),
                OrderPriority = reader.GetString(4),
                OrderDate = reader.GetDateTime(5),
                OrderId = reader.GetInt32(6),
                ShipDate = reader.GetDateTime(7),
                UnitsSold = reader.GetInt32(8),
                UnitPrice = reader.GetDecimal(9),
                UnitCost = reader.GetDecimal(10),
                TotalRevenue = reader.GetDecimal(11),
                TotalCost = reader.GetDecimal(12),
                TotalProfit = reader.GetDecimal(13)
            };
            rows++;
            cells += FixtureData.ExpectedColumns;
            checksum = AddRecord(checksum, record);
        }

        return new Observation(rows, cells, checksum);
    }

    private static Observation ObserveRecords(IEnumerable<SalesRecord> records) {
        int rows = 0;
        int cells = 0;
        long checksum = 17;
        foreach (SalesRecord record in records) {
            rows++;
            cells += FixtureData.ExpectedColumns;
            checksum = AddRecord(checksum, record);
        }

        return new Observation(rows, cells, checksum);
    }

    private static long AddRecord(long checksum, SalesRecord record) {
        checksum = Add(checksum, StringChecksum(record.Region));
        checksum = Add(checksum, StringChecksum(record.Country));
        checksum = Add(checksum, StringChecksum(record.ItemType));
        checksum = Add(checksum, StringChecksum(record.SalesChannel));
        checksum = Add(checksum, StringChecksum(record.OrderPriority));
        checksum = Add(checksum, record.OrderDate.Ticks);
        checksum = Add(checksum, record.OrderId);
        checksum = Add(checksum, record.ShipDate.Ticks);
        checksum = Add(checksum, record.UnitsSold);
        checksum = Add(checksum, DecimalChecksum(record.UnitPrice));
        checksum = Add(checksum, DecimalChecksum(record.UnitCost));
        checksum = Add(checksum, DecimalChecksum(record.TotalRevenue));
        checksum = Add(checksum, DecimalChecksum(record.TotalCost));
        return Add(checksum, DecimalChecksum(record.TotalProfit));
    }

    private static long DecimalChecksum(decimal value) {
        int[] bits = decimal.GetBits(value);
        long checksum = bits[0];
        checksum = Add(checksum, bits[1]);
        checksum = Add(checksum, bits[2]);
        return Add(checksum, bits[3]);
    }

    private static int StringChecksum(string value) {
        int checksum = 17;
        for (int index = 0; index < value.Length; index++) {
            checksum = unchecked((checksum * 31) + value[index]);
        }

        return checksum;
    }

    private static long Add(long checksum, long value) => unchecked((checksum * 31) + value);
}
