#if OFFICEIMO_BENCHMARK_NEW_APIS
using ExcelReader.Core.Reader;
using Sylvan.Data.Excel;
using System.IO.Compression;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Qualifies first and second saves in a separate process from first-use measurement.</summary>
    internal static class NativeXlsbSaveControlQualification {
        internal static void Run() {
            foreach (string shape in new[] { "Mixed", "AllText" })
                foreach (string storage in new[] { "Inline", "Shared" }) {
                    using NativeXlsbSaveControlWorkload workload = new();
                    workload.Prepare(shape, storage);
                    workload.Save();
                    workload.ValidateCurrentOutput($"native-control-{shape}-{storage}-first");
                    workload.Save();
                    workload.ValidateCurrentOutput($"native-control-{shape}-{storage}-second");
                }
            Console.WriteLine("Qualified all four native XLSB controls: first and second saves, every field/type/order, ExcelReader+Sylvan.");
        }

        internal static void Validate(byte[] bytes, int rows, bool allText, bool shared, string lane) {
            if (!allText) {
                NativeWrittenWorkbookValidation.Validate(bytes, rows, ExcelFileFormat.Xlsb, lane, shared);
                using MemoryStream stream = new MemoryStream(bytes, writable: false);
                using IExcelWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
                if (workbook.FirstSheet.Name != "Data" || workbook.FirstSheet.Index != 0)
                    throw new InvalidDataException("The mixed output sheet differs.");
                return;
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using ZipArchive package = new ZipArchive(stream, ZipArchiveMode.Read);
                WrittenWorkbookValidation.ValidateEntries(package);
                WrittenWorkbookValidation.ValidateRelationships(package);
                // Text-only output needs no styles table; validate every declared relationship.
                foreach (string part in new[] { "xl/workbook.bin", "xl/worksheets/sheet1.bin" })
                    if (package.GetEntry(part) == null) throw new InvalidDataException($"The all-text writer omitted {part}.");
                HashSet<string> distinct = new HashSet<string>(TypedWorkbookFixture.Headers, StringComparer.Ordinal);
                for (int index = 1; index <= rows; index++)
                    foreach (string text in NativeXlsbSaveControlWorkload.TextFields(TypedWorkbookFixture.ExpectedRecord(index)))
                        distinct.Add(text);
                XlsbStringStorageValidation.Validate(package, lane, shared, checked((rows + 1) * 4), distinct.Count);
                if (!shared && package.GetEntry("xl/sharedStrings.bin") != null)
                    throw new InvalidDataException("The inline all-text output unexpectedly contains an SST.");
            }

            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using IExcelWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
                if (workbook.SheetCount != 1 || workbook.FirstSheet.Name != "Data" || workbook.FirstSheet.Index != 0)
                    throw new InvalidDataException("The all-text output sheet differs.");
                int index = 0;
                foreach (Row row in workbook.FirstSheet) {
                    if (index > rows || row.ColumnCount != 4) throw new InvalidDataException("The all-text output dimensions differ.");
                    string[] expected = index == 0 ? TypedWorkbookFixture.Headers
                        : NativeXlsbSaveControlWorkload.TextFields(TypedWorkbookFixture.ExpectedRecord(index));
                    for (int column = 0; column < 4; column++)
                        if (row[column].Type != CellType.ExcelString || row[column].GetString() != expected[column])
                            throw new InvalidDataException($"The all-text value/type differs at row {index}, column {column}.");
                    index++;
                }
                if (index != rows + 1) throw new InvalidDataException("The all-text output row count differs.");
            }
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using Sylvan.Data.Excel.ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
                    ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
                if (reader.WorksheetName != "Data" || reader.FieldCount != 4)
                    throw new InvalidDataException("Independent all-text sheet/width differs.");
                for (int column = 0; column < 4; column++)
                    if (reader.GetName(column) != TypedWorkbookFixture.Headers[column])
                        throw new InvalidDataException("Independent all-text header differs.");
                int index = 0;
                while (reader.Read()) {
                    if (++index > rows) throw new InvalidDataException("Independent all-text output has extra rows.");
                    string[] expected = NativeXlsbSaveControlWorkload.TextFields(TypedWorkbookFixture.ExpectedRecord(index));
                    for (int column = 0; column < 4; column++)
                        if (reader.GetExcelDataType(column) != ExcelDataType.String || reader.GetString(column) != expected[column])
                            throw new InvalidDataException($"Independent all-text value/type differs at row {index}, column {column}.");
                }
                if (index != rows || reader.NextResult()) throw new InvalidDataException("Independent all-text row/sheet count differs.");
            }
            BenchmarkInput.WriteWorkbookFixtureIdentity($"validated-output/{lane}/Xlsb/dataRows={rows}", bytes);
            Console.WriteLine($"Validated {lane}: rows={rows + 1}, columns=4, all text, independentReaders=ExcelReader+Sylvan.");
            string? artifacts = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_WRITTEN_ARTIFACTS");
            if (!string.IsNullOrWhiteSpace(artifacts)) {
                if (!Directory.Exists(artifacts))
                    throw new DirectoryNotFoundException("Create the explicit written-artifact directory before qualification.");
                string path = Path.Combine(Path.GetFullPath(artifacts), $"{lane}-Xlsb-{rows}.xlsb");
                File.WriteAllBytes(path, bytes);
                Console.WriteLine($"Retained independently readable native output: {path}.");
            }
        }
    }
}
#endif
