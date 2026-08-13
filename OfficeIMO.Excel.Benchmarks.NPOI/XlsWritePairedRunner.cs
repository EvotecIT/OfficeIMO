using System.Diagnostics;
using System.Globalization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeIMO.Benchmarks;
using OfficeIMO.Excel;

internal static class XlsWritePairedRunner {
    internal static void Run(
        IReadOnlyList<SalesRecord> records,
        int warmupIterations,
        int measuredIterations,
        string? affinityMask,
        string? priorityName) {
        string affinity = affinityMask is null
            ? "unchanged"
            : BenchmarkProcessorAffinity.Apply(affinityMask);
        string priority = priorityName is null
            ? Process.GetCurrentProcess().PriorityClass.ToString()
            : BenchmarkProcessorAffinity.ApplyPriority(priorityName);

        Validate(XlsWriteBenchmarkScenario.WriteOfficeImo(records), records, "OfficeIMO warmup validation");
        Validate(XlsWriteBenchmarkScenario.WriteNpoi(records), records, "NPOI warmup validation");

        for (int index = 0; index < warmupIterations; index++) {
            byte[] officeBytes = XlsWriteBenchmarkScenario.WriteOfficeImo(records);
            byte[] npoiBytes = XlsWriteBenchmarkScenario.WriteNpoi(records);
            Validate(officeBytes, records, $"OfficeIMO warmup {index + 1}");
            Validate(npoiBytes, records, $"NPOI warmup {index + 1}");
        }

        var officeSamples = new double[measuredIterations];
        var npoiSamples = new double[measuredIterations];
        var pairedRatios = new double[measuredIterations];
        int officeBytesWritten = 0;
        int npoiBytesWritten = 0;

        for (int index = 0; index < measuredIterations; index++) {
            TimedWorkbook officeFirst;
            TimedWorkbook officeSecond;
            TimedWorkbook npoiFirst;
            TimedWorkbook npoiSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(() => XlsWriteBenchmarkScenario.WriteOfficeImo(records));
                npoiFirst = Measure(() => XlsWriteBenchmarkScenario.WriteNpoi(records));
                npoiSecond = Measure(() => XlsWriteBenchmarkScenario.WriteNpoi(records));
                officeSecond = Measure(() => XlsWriteBenchmarkScenario.WriteOfficeImo(records));
            } else {
                npoiFirst = Measure(() => XlsWriteBenchmarkScenario.WriteNpoi(records));
                officeFirst = Measure(() => XlsWriteBenchmarkScenario.WriteOfficeImo(records));
                officeSecond = Measure(() => XlsWriteBenchmarkScenario.WriteOfficeImo(records));
                npoiSecond = Measure(() => XlsWriteBenchmarkScenario.WriteNpoi(records));
            }

            Validate(officeFirst.Bytes, records, $"OfficeIMO sample {index + 1} first");
            Validate(officeSecond.Bytes, records, $"OfficeIMO sample {index + 1} second");
            Validate(npoiFirst.Bytes, records, $"NPOI sample {index + 1} first");
            Validate(npoiSecond.Bytes, records, $"NPOI sample {index + 1} second");

            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            npoiSamples[index] = (npoiFirst.Milliseconds + npoiSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / npoiSamples[index];
            officeBytesWritten = officeSecond.Bytes.Length;
            npoiBytesWritten = npoiSecond.Bytes.Length;
        }

        double officeMedian = Median(officeSamples);
        double npoiMedian = Median(npoiSamples);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "Paired XLS write ({0} rows x 5 columns, {1} warmups, {2} ABBA/BAAB samples, affinity {3}, priority {4}): " +
            "OfficeIMO median {5:F3} ms ({6} bytes), NPOI HSSF median {7:F3} ms ({8} bytes), " +
            "ratio of medians {9:F4}, paired ratio median {10:F4} (P25 {11:F4}, P75 {12:F4}).",
            records.Count,
            warmupIterations,
            measuredIterations,
            affinity,
            priority,
            officeMedian,
            officeBytesWritten,
            npoiMedian,
            npoiBytesWritten,
            officeMedian / npoiMedian,
            Median(pairedRatios),
            Percentile(pairedRatios, 0.25d),
            Percentile(pairedRatios, 0.75d)));
    }

    internal static void Validate(byte[] workbookBytes, IReadOnlyList<SalesRecord> records, string source) {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbookBytes);
        string[] expectedHeaders = ["Id", "Region", "Owner", "Amount", "Active"];
        if (reader.FieldCount != expectedHeaders.Length) {
            throw new InvalidDataException($"{source} exposed {reader.FieldCount} fields instead of {expectedHeaders.Length}.");
        }

        for (int column = 0; column < expectedHeaders.Length; column++) {
            if (!string.Equals(reader.GetName(column), expectedHeaders[column], StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"{source} header {column + 1} was '{reader.GetName(column)}' instead of '{expectedHeaders[column]}'.");
            }
        }

        int rowIndex = 0;
        while (reader.Read()) {
            if (rowIndex >= records.Count) {
                throw new InvalidDataException($"{source} emitted more than {records.Count} data rows.");
            }

            SalesRecord expected = records[rowIndex];
            if (reader.GetInt32(0) != expected.Id
                || !string.Equals(reader.GetString(1), expected.Region, StringComparison.Ordinal)
                || !string.Equals(reader.GetString(2), expected.Owner, StringComparison.Ordinal)
                || reader.GetDouble(3) != expected.Amount
                || reader.GetBoolean(4) != expected.Active) {
                throw new InvalidDataException($"{source} row {rowIndex + 2} did not match the input record.");
            }
            rowIndex++;
        }

        if (rowIndex != records.Count) {
            throw new InvalidDataException($"{source} emitted {rowIndex} data rows instead of {records.Count}.");
        }
        if (reader.NextResult()) {
            throw new InvalidDataException($"{source} emitted an unexpected second worksheet.");
        }

        if (source.StartsWith("OfficeIMO", StringComparison.Ordinal)) {
            ValidateWithNpoi(workbookBytes, records, source);
        }
    }

    private static void ValidateWithNpoi(byte[] workbookBytes, IReadOnlyList<SalesRecord> records, string source) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new HSSFWorkbook(stream);
        if (workbook.NumberOfSheets != 1) {
            throw new InvalidDataException($"{source} exposed {workbook.NumberOfSheets} worksheets to NPOI instead of one.");
        }

        ISheet sheet = workbook.GetSheetAt(0);
        string[] expectedHeaders = ["Id", "Region", "Owner", "Amount", "Active"];
        IRow header = sheet.GetRow(0) ?? throw new InvalidDataException($"{source} exposed no header row to NPOI.");
        for (int column = 0; column < expectedHeaders.Length; column++) {
            if (!string.Equals(header.GetCell(column)?.StringCellValue, expectedHeaders[column], StringComparison.Ordinal)) {
                throw new InvalidDataException($"{source} header {column + 1} was not readable through NPOI.");
            }
        }

        for (int index = 0; index < records.Count; index++) {
            SalesRecord expected = records[index];
            IRow row = sheet.GetRow(index + 1) ?? throw new InvalidDataException($"{source} row {index + 2} was not readable through NPOI.");
            if (row.GetCell(0)?.NumericCellValue != expected.Id
                || !string.Equals(row.GetCell(1)?.StringCellValue, expected.Region, StringComparison.Ordinal)
                || !string.Equals(row.GetCell(2)?.StringCellValue, expected.Owner, StringComparison.Ordinal)
                || row.GetCell(3)?.NumericCellValue != expected.Amount
                || row.GetCell(4)?.BooleanCellValue != expected.Active) {
                throw new InvalidDataException($"{source} row {index + 2} did not round-trip through NPOI.");
            }
        }
    }

    private static TimedWorkbook Measure(Func<byte[]> operation) {
        long started = Stopwatch.GetTimestamp();
        byte[] bytes = operation();
        return new TimedWorkbook(Stopwatch.GetElapsedTime(started).TotalMilliseconds, bytes);
    }

    private static double Median(double[] samples) {
        double[] ordered = (double[])samples.Clone();
        Array.Sort(ordered);
        int middle = ordered.Length / 2;
        return (ordered.Length & 1) == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2d
            : ordered[middle];
    }

    private static double Percentile(double[] samples, double percentile) {
        double[] ordered = (double[])samples.Clone();
        Array.Sort(ordered);
        double position = (ordered.Length - 1) * percentile;
        int lower = (int)position;
        int upper = Math.Min(lower + 1, ordered.Length - 1);
        double fraction = position - lower;
        return ordered[lower] + ((ordered[upper] - ordered[lower]) * fraction);
    }

    private readonly record struct TimedWorkbook(double Milliseconds, byte[] Bytes);
}

internal static class XlsWriteBenchmarkScenario {
    internal static byte[] WriteOfficeImo(
        IReadOnlyList<SalesRecord> records,
        Action<string, TimeSpan>? onTiming = null,
        Action<string>? onInfo = null) {
        using ExcelDocument document = ExcelDocument.Create();
        document.Execution.OnTiming = onTiming;
        document.Execution.OnInfo = onInfo;
        ExcelSheet sheet = document.AddWorksheet("Data");
        WriteOfficeImoRows(sheet, records);
        return document.ToBytes(ExcelFileFormat.Xls);
    }

    internal static byte[] WriteNpoi(IReadOnlyList<SalesRecord> records) {
        using var stream = new MemoryStream();
        using var workbook = new HSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Data");
        WriteNpoiRows(sheet, records);
        workbook.Write(stream, leaveOpen: true);
        return stream.ToArray();
    }

    private static void WriteOfficeImoRows(ExcelSheet sheet, IReadOnlyList<SalesRecord> records) {
        sheet.CellValue(1, 1, "Id");
        sheet.CellValue(1, 2, "Region");
        sheet.CellValue(1, 3, "Owner");
        sheet.CellValue(1, 4, "Amount");
        sheet.CellValue(1, 5, "Active");

        for (int index = 0; index < records.Count; index++) {
            int row = index + 2;
            SalesRecord record = records[index];
            sheet.CellValue(row, 1, record.Id);
            sheet.CellValue(row, 2, record.Region);
            sheet.CellValue(row, 3, record.Owner);
            sheet.CellValue(row, 4, record.Amount);
            sheet.CellValue(row, 5, record.Active);
        }
    }

    private static void WriteNpoiRows(ISheet sheet, IReadOnlyList<SalesRecord> records) {
        IRow header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Id");
        header.CreateCell(1).SetCellValue("Region");
        header.CreateCell(2).SetCellValue("Owner");
        header.CreateCell(3).SetCellValue("Amount");
        header.CreateCell(4).SetCellValue("Active");

        for (int index = 0; index < records.Count; index++) {
            IRow row = sheet.CreateRow(index + 1);
            SalesRecord record = records[index];
            row.CreateCell(0).SetCellValue(record.Id);
            row.CreateCell(1).SetCellValue(record.Region);
            row.CreateCell(2).SetCellValue(record.Owner);
            row.CreateCell(3).SetCellValue(record.Amount);
            row.CreateCell(4).SetCellValue(record.Active);
        }
    }
}
