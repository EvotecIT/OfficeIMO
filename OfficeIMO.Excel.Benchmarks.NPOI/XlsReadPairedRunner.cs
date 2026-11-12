using System.Diagnostics;
using System.Globalization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeIMO.Benchmarks;
using OfficeIMO.Excel;

/// <summary>
/// Reads every cell of the first worksheet of existing XLS workbooks through OfficeIMO's data reader and NPOI HSSF.
/// Both engines must produce the same checksum over each cell's position and value (numbers and dates as serial
/// numbers, text as characters) before any timing is reported. Samples alternate ABBA/BAAB order.
/// </summary>
internal static class XlsReadPairedRunner {
    internal static void Run(IReadOnlyList<string> paths, int warmupIterations, int measuredIterations,
        string? affinityMask, string? priorityName) {
        string affinity = affinityMask is null ? "unchanged" : BenchmarkProcessorAffinity.Apply(affinityMask);
        string priority = priorityName is null
            ? Process.GetCurrentProcess().PriorityClass.ToString()
            : BenchmarkProcessorAffinity.ApplyPriority(priorityName);

        foreach (string path in paths) {
            byte[] bytes = File.ReadAllBytes(path);
            Observation office = ReadOfficeImo(bytes);
            Observation npoi = ReadNpoi(bytes);
            if (office != npoi) {
                throw new InvalidDataException(
                    $"{Path.GetFileName(path)}: OfficeIMO read {office} but NPOI read {npoi}; timings withheld.");
            }

            for (int index = 0; index < warmupIterations; index++) {
                Check(ReadOfficeImo(bytes), office, "OfficeIMO warmup");
                Check(ReadNpoi(bytes), office, "NPOI warmup");
            }

            var officeSamples = new double[measuredIterations];
            var npoiSamples = new double[measuredIterations];
            var ratios = new double[measuredIterations];
            for (int index = 0; index < measuredIterations; index++) {
                double o1, o2, n1, n2;
                if ((index & 1) == 0) {
                    o1 = Measure(() => ReadOfficeImo(bytes), office, "OfficeIMO");
                    n1 = Measure(() => ReadNpoi(bytes), office, "NPOI");
                    n2 = Measure(() => ReadNpoi(bytes), office, "NPOI");
                    o2 = Measure(() => ReadOfficeImo(bytes), office, "OfficeIMO");
                } else {
                    n1 = Measure(() => ReadNpoi(bytes), office, "NPOI");
                    o1 = Measure(() => ReadOfficeImo(bytes), office, "OfficeIMO");
                    o2 = Measure(() => ReadOfficeImo(bytes), office, "OfficeIMO");
                    n2 = Measure(() => ReadNpoi(bytes), office, "NPOI");
                }
                officeSamples[index] = (o1 + o2) / 2d;
                npoiSamples[index] = (n1 + n2) / 2d;
                ratios[index] = officeSamples[index] / npoiSamples[index];
            }

            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "Paired XLS read {0} ({1} cells, {2} warmups, {3} ABBA/BAAB samples, affinity {4}, priority {5}): " +
                "OfficeIMO median {6:F3} ms, NPOI HSSF median {7:F3} ms, paired ratio median {8:F4} (P25 {9:F4}, P75 {10:F4}).",
                Path.GetFileName(path), office.Cells, warmupIterations, measuredIterations, affinity, priority,
                Median(officeSamples), Median(npoiSamples), Median(ratios), Percentile(ratios, 0.25), Percentile(ratios, 0.75)));
        }
    }

    /// <summary>Writes the generated four-column workbook that the Excel benchmark project reads as XLS.</summary>
    internal static void WriteGenerated(string path, int rowCount) {
        using var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, path,
            new ExcelWriterOptions { IncludeCellReferences = false, UseSharedStrings = false });
        writer.WriteTable("Data",
            Enumerable.Range(0, rowCount).Select(index => (
                Name: "Name " + index.ToString(CultureInfo.InvariantCulture),
                Id: index + 1,
                CreatedOn: new DateTime(2024, 1, 1).AddDays(index % 365),
                Amount: index * 1.25)),
            new[] { "Name", "Id", "CreatedOn", "Amount" },
            static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
    }

    private static Observation ReadOfficeImo(byte[] bytes) {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(new MemoryStream(bytes, writable: false),
            new ExcelReaderOptions { SheetIndex = 0, HasHeaderRow = false });
        var observation = new Observation();
        int row = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) {
                object value = reader.GetValue(column);
                switch (value) {
                    case DBNull: break;
                    case string text: observation.AddText(row, column, text); break;
                    case DateTime date: observation.AddNumber(row, column, date.ToOADate()); break;
                    case bool flag: observation.AddNumber(row, column, flag ? 1 : 0); break;
                    default: observation.AddNumber(row, column, Convert.ToDouble(value, CultureInfo.InvariantCulture)); break;
                }
            }
            row++;
        }
        return observation;
    }

    private static Observation ReadNpoi(byte[] bytes) {
        using var workbook = new HSSFWorkbook(new MemoryStream(bytes, writable: false));
        ISheet sheet = workbook.GetSheetAt(0);
        var observation = new Observation();
        int firstRow = sheet.FirstRowNum;
        int firstColumn = int.MaxValue;
        for (int rowIndex = firstRow; rowIndex <= sheet.LastRowNum; rowIndex++) {
            IRow? row = sheet.GetRow(rowIndex);
            if (row != null && row.FirstCellNum >= 0) firstColumn = Math.Min(firstColumn, row.FirstCellNum);
        }
        for (int rowIndex = firstRow; rowIndex <= sheet.LastRowNum; rowIndex++) {
            IRow? row = sheet.GetRow(rowIndex);
            if (row == null) continue;
            foreach (ICell cell in row) {
                CellType type = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
                int column = cell.ColumnIndex - firstColumn;
                switch (type) {
                    case CellType.Numeric: observation.AddNumber(rowIndex - firstRow, column, cell.NumericCellValue); break;
                    case CellType.String: observation.AddText(rowIndex - firstRow, column, cell.StringCellValue); break;
                    case CellType.Boolean: observation.AddNumber(rowIndex - firstRow, column, cell.BooleanCellValue ? 1 : 0); break;
                }
            }
        }
        return observation;
    }

    private static double Measure(Func<Observation> read, Observation expected, string engine) {
        long start = Stopwatch.GetTimestamp();
        Observation actual = read();
        double milliseconds = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        Check(actual, expected, engine);
        return milliseconds;
    }

    private static void Check(Observation actual, Observation expected, string engine) {
        if (actual != expected) throw new InvalidDataException($"{engine} read {actual} instead of {expected}.");
    }

    private static double Median(double[] values) => Percentile(values, 0.5);

    private static double Percentile(double[] values, double percentile) {
        double[] sorted = values.OrderBy(value => value).ToArray();
        double position = (sorted.Length - 1) * percentile;
        int lower = (int)Math.Floor(position);
        int upper = (int)Math.Ceiling(position);
        return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
    }

    private record struct Observation(long Cells, ulong Checksum) {
        internal void AddNumber(int row, int column, double value) =>
            Add(row, column, (ulong)BitConverter.DoubleToInt64Bits(value));

        internal void AddText(int row, int column, string text) {
            ulong hash = 1469598103934665603UL;
            foreach (char c in text) hash = (hash ^ c) * 1099511628211UL;
            Add(row, column, hash ^ 0x9E3779B97F4A7C15UL);
        }

        private void Add(int row, int column, ulong value) {
            Cells++;
            unchecked { Checksum += (value ^ ((ulong)row << 20) ^ (ulong)column) * 0xFF51AFD7ED558CCDUL; }
        }
    }
}
