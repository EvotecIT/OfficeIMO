using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using OfficeIMO.Benchmarks;
using OfficeIMO.Data;
using Sylvan.Data.Excel;
using System.Data.Common;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using PeerColumn = ExcelReader.Core.Parser.ExcelColumnAttribute;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The exact fourteen-field class in the upstream real-data typed parsing suite.</summary>
    public sealed class RealDataRecord {
        [PeerColumn("Region"), DataColumn("Region")]
        public string Region { get; set; } = "";
        [PeerColumn("Country"), DataColumn("Country")]
        public string Country { get; set; } = "";
        [PeerColumn("Item Type"), DataColumn("Item Type")]
        public string ItemType { get; set; } = "";
        [PeerColumn("Sales Channel"), DataColumn("Sales Channel")]
        public string SalesChannel { get; set; } = "";
        [PeerColumn("Order Priority"), DataColumn("Order Priority")]
        public string OrderPriority { get; set; } = "";
        [PeerColumn("Order Date"), DataColumn("Order Date")]
        public DateTime OrderDate { get; set; }
        [PeerColumn("Order ID"), DataColumn("Order ID")]
        public long OrderId { get; set; }
        [PeerColumn("Ship Date"), DataColumn("Ship Date")]
        public DateTime ShipDate { get; set; }
        [PeerColumn("Units Sold"), DataColumn("Units Sold")]
        public long UnitsSold { get; set; }
        [PeerColumn("Unit Price"), DataColumn("Unit Price")]
        public double UnitPrice { get; set; }
        [PeerColumn("Unit Cost"), DataColumn("Unit Cost")]
        public double UnitCost { get; set; }
        [PeerColumn("Total Revenue"), DataColumn("Total Revenue")]
        public double TotalRevenue { get; set; }
        [PeerColumn("Total Cost"), DataColumn("Total Cost")]
        public double TotalCost { get; set; }
        [PeerColumn("Total Profit"), DataColumn("Total Profit")]
        public double TotalProfit { get; set; }
    }

    /// <summary>Maps all fourteen columns of the hash-pinned real-data files to materialized classes.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("RealDataTyped")]
    public class RealDataTypedReadBenchmarks {
        private byte[] _bytes = [];
        private long _expected;

        [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb)]
        public ComparisonWorkbookFormat Format { get; set; }

        [GlobalSetup]
        public void Setup() {
            BenchmarkInput.WriteDescription();
            string name = "65K_Records_Data." + Format.ToString().ToLowerInvariant();
            MarkPflug65KFixture.EnsureAuthentic(name);
            _bytes = File.ReadAllBytes(Path.Combine(MarkPflug65KFixture.Root, name));
            _expected = Validate(ReadExcelReader());
            if (Validate(ReadOfficeIMO()) != _expected) throw new InvalidDataException("Typed fixture checksums differ.");
            ExcelReaderAutomatic();
            OfficeIMOAutomatic();
            Console.WriteLine($"Validated fourteen-field typed {Format}: rows={MarkPflug65KFixture.ExpectedRows}, "
                + $"SHA256={MarkPflug65KFixture.GetHashes()[name]}, checksum={_expected}; every property checked independently.");
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderAutomatic() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            long sum = 0;
            int count = 0;
            if (Format == ComparisonWorkbookFormat.Xlsx) {
                using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(stream);
                foreach (RealDataRecord row in ExcelParser.FromAttributes<RealDataRecord>().Parse(workbook.FirstSheet)) {
                    sum = unchecked(sum + Accumulate(row));
                    count++;
                }
            } else {
                using XlsbWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
                foreach (RealDataRecord row in ExcelParser.FromAttributes<RealDataRecord>().Parse(workbook.FirstSheet)) {
                    sum = unchecked(sum + Accumulate(row));
                    count++;
                }
            }
            return Check(sum, count);
        }

        [Benchmark]
        public long OfficeIMOAutomatic() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });
            long sum = 0;
            int count = 0;
            foreach (RealDataRecord row in reader.RowsAs<RealDataRecord>()) {
                sum = unchecked(sum + Accumulate(row));
                count++;
            }
            return Check(sum, count);
        }

        private IEnumerable<RealDataRecord> ReadExcelReader() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            if (Format == ComparisonWorkbookFormat.Xlsx) {
                using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(stream);
                foreach (RealDataRecord row in ExcelParser.FromAttributes<RealDataRecord>().Parse(workbook.FirstSheet)) yield return row;
            } else {
                using XlsbWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
                foreach (RealDataRecord row in ExcelParser.FromAttributes<RealDataRecord>().Parse(workbook.FirstSheet)) yield return row;
            }
        }

        private IEnumerable<RealDataRecord> ReadOfficeIMO() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });
            foreach (RealDataRecord row in reader.RowsAs<RealDataRecord>()) yield return row;
        }

        private long Validate(IEnumerable<RealDataRecord> records) {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using ExcelDataReader oracle = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
                Format == ComparisonWorkbookFormat.Xlsx ? ExcelWorkbookType.ExcelXml : ExcelWorkbookType.ExcelBinary,
                new ExcelDataReaderOptions());
            int count = 0;
            long sum = 0;
            foreach (RealDataRecord row in records) {
                if (!oracle.Read() || oracle.FieldCount != 14) throw new InvalidDataException("Typed row count or width differs.");
                RealDataRecord expected = ReadOracle(oracle);
                object[] actualValues = Values(row), expectedValues = Values(expected);
                for (int column = 0; column < actualValues.Length; column++) {
                    if (!Equals(actualValues[column], expectedValues[column]))
                        throw new InvalidDataException($"Typed real-data row {count + 1}, column {oracle.GetName(column)} differs: "
                            + $"mapped={Describe(actualValues[column])}, independent={Describe(expectedValues[column])}.");
                }
                sum = unchecked(sum + Accumulate(row));
                count++;
            }
            if (count != MarkPflug65KFixture.ExpectedRows || oracle.Read()) throw new InvalidDataException("Incorrect real-data typed row count.");
            return sum;
        }

        private static RealDataRecord ReadOracle(DbDataReader row) => new() {
            Region = row.GetString(0), Country = row.GetString(1), ItemType = row.GetString(2),
            SalesChannel = row.GetString(3), OrderPriority = row.GetString(4), OrderDate = row.GetDateTime(5),
            OrderId = row.GetInt64(6), ShipDate = row.GetDateTime(7), UnitsSold = row.GetInt64(8),
            UnitPrice = row.GetDouble(9), UnitCost = row.GetDouble(10), TotalRevenue = row.GetDouble(11),
            TotalCost = row.GetDouble(12), TotalProfit = row.GetDouble(13),
        };
        private static object[] Values(RealDataRecord row) => [row.Region, row.Country, row.ItemType,
            row.SalesChannel, row.OrderPriority, row.OrderDate, row.OrderId, row.ShipDate, row.UnitsSold,
            row.UnitPrice, row.UnitCost, row.TotalRevenue, row.TotalCost, row.TotalProfit];
        private static string Describe(object value) => value switch {
            DateTime date => date.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            double number => number.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            _ => value.ToString() ?? "null",
        };
        private static long Accumulate(RealDataRecord row) => row.Region.Length + row.Country.Length
            + row.ItemType.Length + row.SalesChannel.Length + row.OrderPriority.Length + row.OrderDate.Ticks
            + row.OrderId + row.ShipDate.Ticks + row.UnitsSold + (long)row.UnitPrice + (long)row.UnitCost
            + (long)row.TotalRevenue + (long)row.TotalCost + (long)row.TotalProfit;
        private long Check(long sum, int count) => sum == _expected && count == MarkPflug65KFixture.ExpectedRows
            ? sum : throw new InvalidDataException("Typed real-data scan returned an incorrect count or checksum.");
    }
}
