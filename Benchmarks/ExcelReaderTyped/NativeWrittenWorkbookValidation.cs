using ExcelReader.Core.Reader;
using OfficeIMO.Excel.LegacyXls.Model;
using Sylvan.Data.Excel;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class NativeWrittenWorkbookValidation {
        internal static void Validate(byte[] bytes, int rowCount, ExcelFileFormat format, string engine,
            bool? expectedSharedStrings = null) {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            if (format == ExcelFileFormat.Xlsb) {
                using MemoryStream zipStream = new MemoryStream(bytes, writable: false);
                using ZipArchive package = new ZipArchive(zipStream, ZipArchiveMode.Read);
                WrittenWorkbookValidation.ValidateEntries(package);
                WrittenWorkbookValidation.ValidateRelationships(package);
                foreach (string part in new[] { "xl/workbook.bin", "xl/styles.bin", "xl/worksheets/sheet1.bin" })
                    if (package.GetEntry(part) == null) throw new InvalidDataException($"The XLSB writer omitted {part}.");
                XlsbStringStorageValidation.Validate(package, rowCount, engine, expectedSharedStrings);
            }
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using IExcelWorkbook workbook = format == ExcelFileFormat.Xls
                    ? ExcelReaderApi.FromXls(stream) : ExcelReaderApi.FromXlsb(stream);
                if (workbook.SheetCount != 1) throw new InvalidDataException("The native writer emitted an incorrect sheet count.");
                int rowIndex = 0;
                foreach (Row row in workbook.FirstSheet) {
                    if (row.ColumnCount != 4) throw new InvalidDataException("The native writer emitted an incorrect row width.");
                    if (rowIndex++ == 0) {
                        for (int column = 0; column < 4; column++)
                            if (row[column].GetString() != TypedWorkbookFixture.Headers[column]) throw new InvalidDataException("The native header differs.");
                    } else {
                        if (row[0].Type != CellType.ExcelString || row[1].Type != CellType.Number || row[2].Type != CellType.Date || row[3].Type != CellType.Number
                            || !row[1].TryParse(CultureInfo.InvariantCulture, out double id) || !row[2].TryGetDateTime(out DateTime date)
                            || !row[3].TryParse(CultureInfo.InvariantCulture, out double value)) throw new InvalidDataException("The native writer lost numeric/date types.");
                        if (id != rowIndex - 1) throw new InvalidDataException("The native writer changed the numeric ID.");
                        TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = row[0].GetString(), Id = checked((int)id), Date = date, Value = value }, rowIndex - 1);
                    }
                }
                if (rowIndex != rowCount + 1) throw new InvalidDataException("The native writer emitted an incorrect row count.");
            }
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using Sylvan.Data.Excel.ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
                    format == ExcelFileFormat.Xls ? ExcelWorkbookType.Excel : ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
                if (reader.FieldCount != 4) throw new InvalidDataException("Independent native reader width differs.");
                for (int column = 0; column < 4; column++)
                    if (reader.GetName(column) != TypedWorkbookFixture.Headers[column]) throw new InvalidDataException("Independent native header differs.");
                int index = 0;
                while (reader.Read()) {
                    if (reader.GetFormat(2)?.Kind != FormatKind.Date) throw new InvalidDataException("The native date format differs.");
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = reader.GetString(0), Id = reader.GetInt32(1),
                        Date = reader.GetDateTime(2), Value = reader.GetDouble(3) }, ++index);
                }
                if (index != rowCount || reader.NextResult()) throw new InvalidDataException("Independent native row/sheet count differs.");
            }
            string structure = "";
            if (format == ExcelFileFormat.Xls) {
                LegacyXlsWorkbook model = LegacyXlsWorkbook.Load(bytes);
                LegacyXlsWorksheetIndex? index = model.Worksheets.Single().RowBlockIndex;
                structure = $", indexRecord={(index == null ? "absent" : "present")}, indexDbCellOffsets={index?.DbCellBlockCount ?? 0}";
            }
            BenchmarkInput.WriteWorkbookFixtureIdentity($"validated-output/{engine}/{format}/dataRows={rowCount}", bytes,
                zipPackage: format == ExcelFileFormat.Xlsb);
            Console.WriteLine($"Validated {engine} {format} cells: rows={rowCount + 1}, columns=4, bytes={bytes.Length}, "
                + $"checksum={TypedWorkbookFixture.ExpectedChecksum(rowCount)}, independentReaders=ExcelReader+Sylvan{structure}.");
            string? artifacts = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_WRITTEN_ARTIFACTS");
            if (!string.IsNullOrWhiteSpace(artifacts)) {
                if (!Directory.Exists(artifacts)) throw new DirectoryNotFoundException("Create the explicit written-artifact directory before qualification.");
                string fileName = $"{engine.Replace(' ', '-')}-{format}-{rowCount}.{format.ToString().ToLowerInvariant()}";
                string path = Path.Combine(Path.GetFullPath(artifacts), fileName);
                File.WriteAllBytes(path, bytes);
                Console.WriteLine($"Retained independently readable native output: {path}.");
            }
        }
    }
}
