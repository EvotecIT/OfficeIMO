#if NET10_0_OR_GREATER
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

public partial class ExcelReaderGeneratedBenchmarks {
    [Benchmark]
    public long ExcelReaderNet_Read() {
        long checksum = 0;
        int count = 0;
        switch (Format) {
            case ExcelFileFormat.Xlsx:
                using (var reader = ExcelReaderApi.FromXlsxFile(_path)) {
                    foreach (Row row in reader.FirstSheet) ReadRow(row, reader.IsDate1904, ref checksum, ref count);
                }
                break;
            case ExcelFileFormat.Xlsb:
                using (var reader = ExcelReaderApi.FromXlsbFile(_path)) {
                    foreach (Row row in reader.FirstSheet) ReadRow(row, reader.IsDate1904, ref checksum, ref count);
                }
                break;
            case ExcelFileFormat.Xls:
                using (var reader = ExcelReaderApi.FromXlsFile(_path)) {
                    foreach (Row row in reader.FirstSheet) ReadRow(row, reader.IsDate1904, ref checksum, ref count);
                }
                break;
            default: throw new ArgumentOutOfRangeException();
        }
        return Check(checksum, count - 1);
    }

    private void ReadRow(Row row, bool date1904, ref long checksum, ref int count) {
        if (count++ == 0) {
            for (int i = 0; i < Headers.Length; i++)
                if (row[i].GetString() != Headers[i]) throw new InvalidDataException("Incorrect header.");
            return;
        }
        string name = row[0].GetString();
        if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id)
            || !row[2].TryGetDateTime(date1904, out DateTime date)
            || !row[3].TryParse(CultureInfo.InvariantCulture, out double amount)) throw new InvalidDataException("Incorrect typed value.");
        if (_validate) Validate(count - 2, name, id, date, amount);
        checksum = Add(checksum, name, id, date, amount);
    }
}
#endif
