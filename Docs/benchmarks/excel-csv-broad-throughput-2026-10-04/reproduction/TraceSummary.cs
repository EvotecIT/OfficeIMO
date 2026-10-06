using Microsoft.Diagnostics.Tracing.Etlx;
using Microsoft.Diagnostics.Tracing.Parsers.Clr;

static class TraceSummary {
    public static void Run(string path) {
        using var log = new TraceLog(TraceLog.CreateFromEventPipeDataFile(path));
        var types = new Dictionary<string, long>();
        var owners = new Dictionary<string, long>();
        var methods = new Dictionary<string, long>();
        var cpuLeaves = new Dictionary<string, long>();
        var cpuInclusive = new Dictionary<string, long>();
        long cpuSamples = 0;
        long bytes = 0, samples = 0;
        foreach (var data in log.Events) {
            bool cpu = data.ProviderName == "Microsoft-DotNETCore-SampleProfiler";
            if (!cpu && data is not GCAllocationTickTraceData) continue;
            var frames = new List<string>();
            for (var stack = data.CallStack(); stack != null; stack = stack.Caller)
                frames.Add(stack.CodeAddress.FullMethodName);
            string root = Environment.GetEnvironmentVariable("OFFICEIMO_TRACE_ROOT") ?? "CopyWorksheetFromPackage";
            if (!frames.Any(frame => frame.Contains(root, StringComparison.Ordinal))) continue;
            if (cpu) {
                cpuSamples++;
                string leaf = frames.FirstOrDefault() ?? "unknown";
                cpuLeaves[leaf] = cpuLeaves.GetValueOrDefault(leaf) + 1;
                foreach (string frame in frames.Distinct()) cpuInclusive[frame] = cpuInclusive.GetValueOrDefault(frame) + 1;
                continue;
            }
            var allocation = (GCAllocationTickTraceData)data;
            long amount = allocation.AllocationAmount64;
            bytes += amount;
            samples++;
            types[allocation.TypeName] = types.GetValueOrDefault(allocation.TypeName) + amount;
            string owner = frames.FirstOrDefault(frame => frame.StartsWith("OfficeIMO.", StringComparison.Ordinal)) ?? "unknown";
            owners[owner] = owners.GetValueOrDefault(owner) + amount;
            foreach (string frame in frames.Distinct()) methods[frame] = methods.GetValueOrDefault(frame) + amount;
        }
        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new {
            AllocationSamples = samples, SampledAllocationBytes = bytes,
            CpuSamples = cpuSamples,
            CpuLeaves = cpuLeaves.OrderByDescending(pair => pair.Value).Take(35),
            CpuInclusive = cpuInclusive.OrderByDescending(pair => pair.Value).Take(45),
            OfficeFrames = cpuInclusive.Where(pair => pair.Key.Contains("OfficeIMO", StringComparison.Ordinal)).OrderByDescending(pair => pair.Value).Take(35),
            Types = types.OrderByDescending(pair => pair.Value).Take(20),
            FirstOfficeFrame = owners.OrderByDescending(pair => pair.Value).Take(25),
            InclusiveFrames = methods.OrderByDescending(pair => pair.Value).Take(45)
        }, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
    }
}
