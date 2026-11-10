using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;

// Usage: ExcelEditableMemory <read|edit|write> <iterations> [workbook.xlsx]
//   read  - load the workbook and enumerate every cell of its first sheet
//   edit  - load the workbook, add one cell below its data and save it to memory
//   write - create 100,000 rows of mixed values through the editable model and save them
// Run one mode per process: the peak working set then belongs to that workload alone.
internal static class Program {
    private const int WriteRows = 100_000;

    private static int Main(string[] args) {
        if (args.Length < 2) {
            Console.Error.WriteLine("Usage: ExcelEditableMemory <read|edit|write> <iterations> [workbook.xlsx]");
            return 2;
        }

        string mode = args[0];
        int iterations = int.Parse(args[1], CultureInfo.InvariantCulture);
        byte[] input = mode == "write" ? Array.Empty<byte>() : File.ReadAllBytes(args[2]);
        Func<long> run = mode switch {
            "read" => () => Read(input),
            "edit" => () => Edit(input),
            "write" => Write,
            _ => throw new ArgumentException("Unknown mode " + mode)
        };

        long check = run();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long baseline = GC.GetTotalMemory(forceFullCollection: true);
        long retained = Retained(mode, input) - baseline;

        var times = new List<double>(iterations);
        long allocatedBefore = AllocatedBytes();
        for (int iteration = 0; iteration < iterations; iteration++) {
            var watch = Stopwatch.StartNew();
            if (run() != check) throw new InvalidDataException("The workload returned a different result.");
            times.Add(watch.Elapsed.TotalMilliseconds);
        }
        double allocated = (AllocatedBytes() - allocatedBefore) / (double)iterations / 1048576;

        times.Sort();
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
            "{0} {1}: median {2:F1} ms, min {3:F1} ms, allocated {4:F1} MB/op, retained {5:F1} MB, peak working set {6:F1} MB, check {7}",
            Runtime(), mode, times[times.Count / 2], times[0], allocated, retained / 1048576.0, process.PeakWorkingSet64 / 1048576.0, check));
        return 0;
    }

    private static long Read(byte[] input) {
        long checksum = 0;
        using var document = ExcelDocument.Load(new MemoryStream(input, writable: false));
        foreach (ExcelCellValueInfo cell in document.Sheets[0].EnumerateCells()) {
            checksum += (cell.Value?.ToString() ?? string.Empty).Length;
        }
        return checksum;
    }

    private static long Edit(byte[] input) {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(new MemoryStream(input, writable: false))) {
            ExcelSheet sheet = document.Sheets[0];
            int lastRow = sheet.EnumerateCells().Select(cell => cell.Row).DefaultIfEmpty(0).Max();
            sheet.Cell(lastRow + 1, 1).SetValue("Edited");
            document.Save(output);
        }
        using var reloaded = ExcelDocument.Load(new MemoryStream(output.ToArray(), writable: false));
        return reloaded.Sheets[0].EnumerateCells().LongCount();
    }

    private static long Write() {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create(output)) {
            Populate(document);
            document.Save(output);
        }
        using var reloaded = ExcelDocument.Load(new MemoryStream(output.ToArray(), writable: false));
        return reloaded.Sheets[0].EnumerateCells().LongCount();
    }

    private static void Populate(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        var start = new DateTime(2026, 1, 1);
        for (int row = 1; row <= WriteRows; row++) {
            sheet.Cell(row, 1).SetValue("Item " + row.ToString(CultureInfo.InvariantCulture));
            sheet.Cell(row, 2).SetValue(row * 1.25);
            sheet.Cell(row, 3).SetValue(start.AddDays(row % 365));
            sheet.Cell(row, 4).SetValue(row % 2 == 0);
            sheet.Cell(row, 5).SetValue(row % 7 == 0 ? "North" : "South");
            sheet.Cell(row, 6).SetValue(row);
        }
    }

    /// <summary>Managed heap held by an open workbook: loaded and fully read, or populated and not yet saved.</summary>
    private static long Retained(string mode, byte[] input) {
        if (mode == "write") {
            using var created = ExcelDocument.Create(new MemoryStream());
            Populate(created);
            long heldWrite = GC.GetTotalMemory(forceFullCollection: true);
            GC.KeepAlive(created);
            return heldWrite;
        }

        using var document = ExcelDocument.Load(new MemoryStream(input, writable: false));
        foreach (ExcelCellValueInfo _ in document.Sheets[0].EnumerateCells()) { }
        long held = GC.GetTotalMemory(forceFullCollection: true);
        GC.KeepAlive(document);
        return held;
    }

    private static long AllocatedBytes() {
#if NETFRAMEWORK
        AppDomain.MonitoringIsEnabled = true;
        return AppDomain.CurrentDomain.MonitoringTotalAllocatedMemorySize;
#else
        return GC.GetTotalAllocatedBytes(precise: true);
#endif
    }

    private static string Runtime() =>
#if NETFRAMEWORK
        "net472";
#else
        "net10.0";
#endif
}
