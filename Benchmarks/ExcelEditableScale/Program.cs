using System.Diagnostics;
using System.Globalization;

// Usage: ExcelEditableScale <library> <lane> <iterations> [workbook.xlsx]
//   library: officeimo, closedxml, epplus, npoi, xlibur, sdk
//   lane:    write1m, styles, merged, cfdv, namestables, recalc, autofit, preserve (needs the workbook)
// Runs one library and one lane per process and prints the median and minimum time, allocation per run and the
// process peak working set, then validates the output. autofit times only the autofit call; recalc times loading the
// generated formula workbook, changing its lookup values, recalculating and reading the results.
// "sdk recalc-fixture <path>" writes the formula workbook the recalc lane opens.
internal static class Program {
    private static int Main(string[] args) {
        if (args.Length == 3 && args[0] == "sdk" && args[1] == "recalc-fixture") {
            File.WriteAllBytes(args[2], SdkLanes.RecalcFixture());
            return 0;
        }
        if (args.Length < 3 || !int.TryParse(args[2], NumberStyles.None, CultureInfo.InvariantCulture, out int iterations) || iterations < 1) {
            Console.Error.WriteLine("Usage: ExcelEditableScale <library> <lane> <iterations> [workbook.xlsx]");
            return 2;
        }

        string library = args[0];
        string lane = args[1];
        byte[]? input = args.Length > 3 ? File.ReadAllBytes(args[3]) : null;
        if ((lane is "preserve" or "recalc") && input == null) {
            Console.Error.WriteLine(lane + " needs a workbook path.");
            return 2;
        }
        Func<(double Milliseconds, byte[]? Output, double Check)>? run = Resolve(library, lane, input);
        if (run == null) {
            Console.WriteLine($"{library} {lane}: not applicable");
            return 0;
        }

        var times = new List<double>(iterations);
        byte[]? output = null;
        double check = 0;
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        for (int iteration = 0; iteration < iterations; iteration++) {
            (double milliseconds, output, check) = run();
            times.Add(milliseconds);
        }
        double allocated = (GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore) / (double)iterations / 1048576;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        double peak = process.PeakWorkingSet64 / 1048576.0;

        string? dump = Environment.GetEnvironmentVariable("EXCEL_SCALE_DUMP");
        if (!string.IsNullOrEmpty(dump) && output != null) File.WriteAllBytes(dump, output);

        string validation;
        try {
            if (lane == "recalc") {
                double expected = Workloads.ExpectedRecalcChecksum();
                validation = Math.Abs(check - expected) <= Math.Abs(expected) * 1e-9
                    ? "total " + check.ToString("R", CultureInfo.InvariantCulture)
                    : throw new InvalidDataException($"Validation failed: total {check.ToString("R", CultureInfo.InvariantCulture)}, expected {expected.ToString("R", CultureInfo.InvariantCulture)}");
            } else {
                validation = Validation.Check(lane, output!, input);
            }
        } catch (Exception exception) {
            validation = "INVALID " + exception.Message;
        }

        times.Sort();
        Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
            "{0} {1}: median {2:F1} ms, min {3:F1} ms, allocated {4:F1} MB/op, peak working set {5:F1} MB, output {6:F1} MB, {7}",
            library, lane, times[times.Count / 2], times[0], allocated, peak, (output?.Length ?? 0) / 1048576.0, validation));
        return validation.StartsWith("INVALID", StringComparison.Ordinal) ? 1 : 0;
    }

    private static Func<(double, byte[]?, double)>? Resolve(string library, string lane, byte[]? input) {
        Func<(double, byte[]?, double)> Timed(Func<byte[]> work) => () => {
            var watch = Stopwatch.StartNew();
            byte[] output = work();
            return (watch.Elapsed.TotalMilliseconds, output, 0);
        };
        Func<(double, byte[]?, double)> TimedCheck(Func<double> work) => () => {
            var watch = Stopwatch.StartNew();
            double check = work();
            return (watch.Elapsed.TotalMilliseconds, null, check);
        };
        Func<(double, byte[]?, double)> Inner(Func<(double, byte[])> work) => () => {
            (double milliseconds, byte[] output) = work();
            return (milliseconds, output, 0);
        };

        return (library, lane) switch {
            ("officeimo", "write1m") => Timed(OfficeImoLanes.Write1M),
            ("officeimo", "styles") => Timed(OfficeImoLanes.Styles),
            ("officeimo", "merged") => Timed(OfficeImoLanes.Merged),
            ("officeimo", "cfdv") => Timed(OfficeImoLanes.CfDv),
            ("officeimo", "namestables") => Timed(OfficeImoLanes.NamesTables),
            ("officeimo", "preserve") => Timed(() => OfficeImoLanes.Preserve(input!)),
            ("officeimo", "recalc") => TimedCheck(() => OfficeImoLanes.Recalc(input!)),
            ("officeimo", "autofit") => Inner(OfficeImoLanes.AutoFit),

            ("closedxml", "write1m") => Timed(ClosedXmlLanes.Write1M),
            ("closedxml", "styles") => Timed(ClosedXmlLanes.Styles),
            ("closedxml", "merged") => Timed(ClosedXmlLanes.Merged),
            ("closedxml", "cfdv") => Timed(ClosedXmlLanes.CfDv),
            ("closedxml", "namestables") => Timed(ClosedXmlLanes.NamesTables),
            ("closedxml", "preserve") => Timed(() => ClosedXmlLanes.Preserve(input!)),
            ("closedxml", "recalc") => TimedCheck(() => ClosedXmlLanes.Recalc(input!)),
            ("closedxml", "autofit") => Inner(ClosedXmlLanes.AutoFit),

            ("xlibur", "write1m") => Timed(XLiburLanes.Write1M),
            ("xlibur", "styles") => Timed(XLiburLanes.Styles),
            ("xlibur", "merged") => Timed(XLiburLanes.Merged),
            ("xlibur", "cfdv") => Timed(XLiburLanes.CfDv),
            ("xlibur", "namestables") => Timed(XLiburLanes.NamesTables),
            ("xlibur", "preserve") => Timed(() => XLiburLanes.Preserve(input!)),
            ("xlibur", "recalc") => TimedCheck(() => XLiburLanes.Recalc(input!)),
            ("xlibur", "autofit") => Inner(XLiburLanes.AutoFit),

            ("epplus", "write1m") => Timed(EpPlusLanes.Write1M),
            ("epplus", "styles") => Timed(EpPlusLanes.Styles),
            ("epplus", "merged") => Timed(EpPlusLanes.Merged),
            ("epplus", "cfdv") => Timed(EpPlusLanes.CfDv),
            ("epplus", "namestables") => Timed(EpPlusLanes.NamesTables),
            ("epplus", "preserve") => Timed(() => EpPlusLanes.Preserve(input!)),
            ("epplus", "recalc") => TimedCheck(() => EpPlusLanes.Recalc(input!)),
            ("epplus", "autofit") => Inner(EpPlusLanes.AutoFit),

            ("npoi", "write1m") => Timed(NpoiLanes.Write1M),
            ("npoi", "styles") => Timed(NpoiLanes.Styles),
            ("npoi", "merged") => Timed(NpoiLanes.Merged),
            ("npoi", "cfdv") => Timed(NpoiLanes.CfDv),
            ("npoi", "namestables") => Timed(NpoiLanes.NamesTables),
            ("npoi", "preserve") => Timed(() => NpoiLanes.Preserve(input!)),
            ("npoi", "recalc") => TimedCheck(() => NpoiLanes.Recalc(input!)),
            ("npoi", "autofit") => Inner(NpoiLanes.AutoFit),

            ("sdk", "write1m") => Timed(SdkLanes.Write1M),
            ("sdk", "preserve") => Timed(() => SdkLanes.Preserve(input!)),
            _ => null
        };
    }
}
