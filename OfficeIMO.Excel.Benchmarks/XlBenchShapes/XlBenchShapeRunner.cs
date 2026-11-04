using System.Diagnostics.Tracing;
using System.Globalization;
using OfficeIMO.Benchmarks;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Opt-in OfficeIMO and plain Open XML SDK comparison of the public XLBench scenario shapes.</summary>
internal static class XlBenchShapeRunner {
    private sealed record Scenario(string Name, Func<int> OfficeImo, Func<int>? OpenXml, Action<string, byte[]>? Validate, Func<byte[]>? OfficeImoOutput, Func<byte[]>? OpenXmlOutput);

    private static readonly Scenario[] Scenarios = [
        new("write", () => OfficeWork(XlBenchShapeScenarios.OfficeImoWrite()), () => OfficeWork(XlBenchShapeScenarios.OpenXmlWrite()),
            XlBenchShapeValidation.Write, XlBenchShapeScenarios.OfficeImoWrite, XlBenchShapeScenarios.OpenXmlWrite),
        new("read-all", () => (int)XlBenchShapeScenarios.OfficeImoReadAll(), () => (int)XlBenchShapeScenarios.OpenXmlReadAll(), null, null, null),
        new("sheets-20", () => (int)XlBenchShapeScenarios.OfficeImoSheets(), () => (int)XlBenchShapeScenarios.OpenXmlSheets(), null, null, null),
        new("read-while-writing", () => (int)XlBenchShapeScenarios.OfficeImoReadWhileWriting(), () => (int)XlBenchShapeScenarios.OpenXmlReadWhileWriting(), null, null, null),
        new("wide-read", () => (int)XlBenchShapeScenarios.OfficeImoWideRead(), () => (int)XlBenchShapeScenarios.OpenXmlWideRead(), null, null, null),
        new("edited-read", () => (int)XlBenchShapeScenarios.OfficeImoEditedRead(), () => (int)XlBenchShapeScenarios.OpenXmlEditedRead(), null, null, null),
        new("edited-sheets-read", () => (int)XlBenchShapeScenarios.OfficeImoEditedSheetsRead(), () => (int)XlBenchShapeScenarios.OpenXmlEditedSheetsRead(), null, null, null),
        new("report", () => OfficeWork(XlBenchShapeScenarios.OfficeImoReport()), () => OfficeWork(XlBenchShapeScenarios.OpenXmlReport()),
            XlBenchShapeValidation.Report, XlBenchShapeScenarios.OfficeImoReport, XlBenchShapeScenarios.OpenXmlReport),
        new("edit", () => (int)XlBenchShapeScenarios.OfficeImoEdit(null), () => (int)XlBenchShapeScenarios.OpenXmlEdit(null),
            XlBenchShapeValidation.Edit, () => Capture(XlBenchShapeScenarios.OfficeImoEdit), () => Capture(XlBenchShapeScenarios.OpenXmlEdit)),
        new("edit-save", () => OfficeWork(Capture(XlBenchShapeScenarios.OfficeImoEdit)), () => OfficeWork(Capture(XlBenchShapeScenarios.OpenXmlEdit)), null, null, null),
        new("insert", () => (int)XlBenchShapeScenarios.OfficeImoInsert(null), () => (int)XlBenchShapeScenarios.OpenXmlInsert(null),
            XlBenchShapeValidation.Insert, () => Capture(XlBenchShapeScenarios.OfficeImoInsert), () => Capture(XlBenchShapeScenarios.OpenXmlInsert)),
        new("insert-save", () => OfficeWork(Capture(XlBenchShapeScenarios.OfficeImoInsert)), () => OfficeWork(Capture(XlBenchShapeScenarios.OpenXmlInsert)), null, null, null),
        new("calc", () => (int)XlBenchShapeScenarios.OfficeImoCalc(null), null,
            XlBenchShapeValidation.Calc, () => Capture(XlBenchShapeScenarios.OfficeImoCalc), null),
        new("template", () => OfficeWork(XlBenchShapeScenarios.OfficeImoTemplate()), () => OfficeWork(XlBenchShapeScenarios.OpenXmlTemplate()),
            XlBenchShapeValidation.Template, XlBenchShapeScenarios.OfficeImoTemplate, XlBenchShapeScenarios.OpenXmlTemplate),
        new("properties", () => OfficeWork(XlBenchShapeScenarios.OfficeImoProperties()), () => OfficeWork(XlBenchShapeScenarios.OpenXmlProperties()),
            XlBenchShapeValidation.Properties, XlBenchShapeScenarios.OfficeImoProperties, XlBenchShapeScenarios.OpenXmlProperties)
    ];

    /// <summary>Arguments: [iterations] [warmup] [comma-separated scenarios]. Each pair is validated before timing.</summary>
    internal static void Run(string[] arguments) {
        int iterations = arguments.Length > 1 ? int.Parse(arguments[1], CultureInfo.InvariantCulture) : 9;
        int warmup = arguments.Length > 2 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 2;
        if (iterations <= 0 || warmup < 0) {
            throw new ArgumentOutOfRangeException(nameof(arguments));
        }

        Console.WriteLine("| Scenario | OfficeIMO ms | OfficeIMO MB | SDK ms | SDK MB | Time ratio | Alloc ratio | Output |");
        Console.WriteLine("|---|---:|---:|---:|---:|---:|---:|---|");
        foreach (Scenario scenario in Select(arguments.Length > 3 ? arguments[3] : null)) {
            string output = "n/a";
            if (scenario.Validate != null) {
                scenario.Validate($"{scenario.Name} OfficeIMO", scenario.OfficeImoOutput!());
                if (scenario.OpenXmlOutput != null) {
                    scenario.Validate($"{scenario.Name} OpenXML", scenario.OpenXmlOutput());
                }
                output = "validated";
            }

            if (scenario.OpenXml == null) {
                // The SDK has no calculation engine, so the scenario has no SDK counterpart.
                var alone = BenchmarkMeasurement.MeasureGroup(warmup, iterations, [scenario.OfficeImo])[0];
                Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                    $"| {scenario.Name} | {alone.MedianMilliseconds:F1} | {alone.MedianAllocatedBytes / 1_048_576d:F2} | n/a | n/a | n/a | n/a | {output} |"));
                continue;
            }

            var results = BenchmarkMeasurement.MeasureGroup(warmup, iterations, [scenario.OfficeImo, scenario.OpenXml]);
            var office = results[0];
            var sdk = results[1];
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"| {scenario.Name} | {office.MedianMilliseconds:F1} | {office.MedianAllocatedBytes / 1_048_576d:F2} | {sdk.MedianMilliseconds:F1} | {sdk.MedianAllocatedBytes / 1_048_576d:F2} | {office.MedianMilliseconds / sdk.MedianMilliseconds:F2} | {office.MedianAllocatedBytes / sdk.MedianAllocatedBytes:F2} | {output} |"));
        }
    }

    /// <summary>Arguments: scenario [iterations]. Repeats OfficeIMO alone and prints sampled allocation by type.</summary>
    internal static void Profile(string[] arguments) {
        Scenario scenario = Select(arguments.Length > 1 ? arguments[1] : throw new ArgumentException("A scenario is required.")).Single();
        int iterations = arguments.Length > 2 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 10;
        scenario.OfficeImo();
        using var listener = new AllocationListener();
        long before = GC.GetAllocatedBytesForCurrentThread();
        for (int i = 0; i < iterations; i++) {
            scenario.OfficeImo();
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Thread.Sleep(500);
        Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"{scenario.Name}: {allocated / 1_048_576d / iterations:F2} MB per iteration"));
        foreach (var (type, share) in listener.TopTypes(30)) {
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"{share,6:P1} {type}"));
        }
    }

    private static IEnumerable<Scenario> Select(string? filter) {
        if (filter == null) {
            return Scenarios;
        }
        var names = filter.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return names.Select(name => Array.Find(Scenarios, scenario => string.Equals(scenario.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"Unknown scenario '{name}'. Known: {string.Join(", ", Scenarios.Select(static scenario => scenario.Name))}."));
    }

    private static int OfficeWork(byte[] output) => output.Length;

    private static byte[] Capture(Func<Stream?, double> operation) {
        using var output = new MemoryStream();
        operation(output);
        return output.ToArray();
    }

    /// <summary>Samples the runtime's allocation ticks. It has no stacks, so read it alongside phase timing.</summary>
    private sealed class AllocationListener : EventListener {
        private readonly Dictionary<string, long> _bytes = new(StringComparer.Ordinal);

        protected override void OnEventSourceCreated(EventSource eventSource) {
            if (eventSource.Name == "Microsoft-Windows-DotNETRuntime") {
                EnableEvents(eventSource, EventLevel.Verbose, (EventKeywords)1);
            }
        }

        protected override void OnEventWritten(EventWrittenEventArgs eventData) {
            if (eventData.EventName == null || !eventData.EventName.StartsWith("GCAllocationTick", StringComparison.Ordinal) || eventData.Payload == null) {
                return;
            }
            int typeIndex = eventData.PayloadNames!.IndexOf("TypeName");
            int amountIndex = eventData.PayloadNames.IndexOf("AllocationAmount64");
            if (typeIndex < 0 || amountIndex < 0) {
                return;
            }
            string type = eventData.Payload[typeIndex]?.ToString() ?? "unknown";
            long amount = Convert.ToInt64(eventData.Payload[amountIndex], CultureInfo.InvariantCulture);
            lock (_bytes) {
                _bytes[type] = _bytes.GetValueOrDefault(type) + amount;
            }
        }

        internal IEnumerable<(string Type, double Share)> TopTypes(int count) {
            lock (_bytes) {
                double total = Math.Max(1, _bytes.Values.Sum());
                return _bytes.OrderByDescending(static pair => pair.Value).Take(count).Select(pair => (pair.Key, pair.Value / total)).ToArray();
            }
        }
    }
}
