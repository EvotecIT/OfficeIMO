using System.Diagnostics;
using System.Globalization;
using BenchmarkDotNet.Running;
using OfficeIMO.XRayComparison;

// Commands:
//   cases                         list every fixture with size and origin
//   validate [--lanes a,b]        run each workload once per library, score and save outputs
//   peak <lane> <case> <side>     (child) run one operation 3x in a fresh process and report peak working set
//   peaks [--cases a,b]           spawn `peak` for each large case and library
//   loop <lane> <case> <side> <n> run one operation n times with timing, for a profiler
//   anything else                 BenchmarkDotNet switcher (--filter, --list, ...)
if (args.Length > 0) {
    switch (args[0].ToLowerInvariant()) {
        case "cases":
            foreach (BenchCase item in Corpus.Extraction) {
                Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"{item.Id,-8} {item.Bytes,12:N0}  {item.Origin,-40} {Path.GetRelativePath(Corpus.RepoRoot, item.Path)}"));
            }
            return 0;
        case "validate":
            return Validation.Run(args.Skip(1).ToArray());
        case "peak":
            return PeakProbe.Child(args[1], args[2], args[3]);
        case "loop":
            return PeakProbe.Loop(args[1], args[2], args[3], int.Parse(args[4], CultureInfo.InvariantCulture));
        case "peaks":
            return PeakProbe.Driver(args.Skip(1).ToArray());
    }
}

BenchmarkSwitcher.FromAssembly(typeof(Corpus).Assembly).Run(args);
return 0;

/// <summary>Process-isolated peak working set: each library and case runs in its own fresh process.</summary>
internal static class PeakProbe {
    private static readonly string[] Lanes = ["text", "markdown", "pdf"];

    private static Func<object> Operation(string lane, string caseId, string side) {
        BenchCase item = Corpus.Get(caseId);
        byte[] bytes = File.ReadAllBytes(item.Path);
        bool office = side == "oi";
        return lane switch {
            "text" => office ? () => Workloads.OfficeImoText(bytes, item) : () => Workloads.XRayText(bytes, item),
            "markdown" => office ? () => Workloads.OfficeImoMarkdown(bytes, item) : () => Workloads.XRayMarkdown(bytes, item),
            "pdf" => office ? () => Workloads.OfficeImoPdf(item) : () => Workloads.XRayPdf(item),
            _ => throw new ArgumentException(lane)
        };
    }

    /// <summary>Runs one operation repeatedly with a per-run timing line, for attaching a profiler.</summary>
    public static int Loop(string lane, string caseId, string side, int count) {
        Func<object> run = Operation(lane, caseId, side);
        for (int index = 0; index < count; index++) {
            long allocated = GC.GetAllocatedBytesForCurrentThread();
            var watch = Stopwatch.StartNew();
            GC.KeepAlive(run());
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"{index}	{watch.Elapsed.TotalMilliseconds:0.0} ms	{(GC.GetAllocatedBytesForCurrentThread() - allocated) / 1048576.0:0.0} MB"));
        }
        return 0;
    }

    public static int Child(string lane, string caseId, string side) {
        Func<object> run = Operation(lane, caseId, side);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long before = process.WorkingSet64;
        var watch = Stopwatch.StartNew();
        object? keep = null;
        for (int index = 0; index < 3; index++) keep = run();
        watch.Stop();
        GC.KeepAlive(keep);
        process.Refresh();
        Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
            $"{lane}\t{caseId}\t{side}\t{process.PeakWorkingSet64 / 1048576.0:0.0}\t{before / 1048576.0:0.0}\t{watch.Elapsed.TotalMilliseconds / 3:0.0}"));
        return 0;
    }

    public static int Driver(string[] args) {
        string[] cases = (Validation.Option(args, "--cases") ?? "docx-l,xlsx-l,pptx-l,pdf-l").Split(',');
        string[] lanes = (Validation.Option(args, "--lanes") ?? string.Join(',', Lanes)).Split(',');
        Console.WriteLine("lane\tcase\tside\tpeakWorkingSetMB\tstartWorkingSetMB\tcoldMsPerOp(3 runs, includes JIT)");
        string self = Environment.ProcessPath!;
        foreach (string lane in lanes) {
            foreach (string caseId in cases) {
                BenchCase item = Corpus.Get(caseId);
                bool applies = lane switch {
                    "text" or "markdown" => item.Format is "docx" or "xlsx" or "pptx" or "pdf",
                    "pdf" => item.Format is "docx" or "xlsx" or "pptx",
                    _ => false
                };
                if (!applies) continue;
                foreach (string side in new[] { "oi", "xr" }) {
                    var start = new ProcessStartInfo(self) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
                    foreach (string argument in new[] { "peak", lane, caseId, side }) start.ArgumentList.Add(argument);
                    using Process child = Process.Start(start)!;
                    string output = child.StandardOutput.ReadToEnd();
                    string error = child.StandardError.ReadToEnd();
                    child.WaitForExit();
                    Console.Write(child.ExitCode == 0 ? output : $"{lane}\t{caseId}\t{side}\tfailed: {error.Split('\n')[0]}\n");
                }
            }
        }
        return 0;
    }
}
