using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Loader;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

if (args.Length > 0 && args[0] == "--xml-compare") {
    foreach (string scenario in args.Skip(1)) {
        byte[] Export(string version) {
            var (instance, _) = Snapshot.Load(scenario, version);
            object table = instance.GetType().GetField("_table", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(instance)!;
            if (scenario.StartsWith("ExcelDefault", StringComparison.Ordinal))
                return (byte[])instance.GetType().GetMethod("WritePackage", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(instance, null)!;
            Type runner = instance.GetType().Assembly.GetType("OfficeIMO.Excel.Benchmarks.ExcelLibraryComparisonRunner")!;
            return (byte[])runner.GetMethod("OfficeImoWriteDataReaderCompactPackageBytes", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, [table])!;
        }
        byte[] before = Export("baseline"), after = Export("candidate");
        using var oldZip = new System.IO.Compression.ZipArchive(new MemoryStream(before));
        using var newZip = new System.IO.Compression.ZipArchive(new MemoryStream(after));
        var hashes = new List<object>();
        foreach (var oldEntry in oldZip.Entries) {
            var newEntry = newZip.GetEntry(oldEntry.FullName) ?? throw new InvalidOperationException("Package part missing.");
            using var oldStream = oldEntry.Open();
            using var newStream = newEntry.Open();
            byte[] oldHash = System.Security.Cryptography.SHA256.HashData(oldStream);
            byte[] newHash = System.Security.Cryptography.SHA256.HashData(newStream);
            bool equal = oldHash.AsSpan().SequenceEqual(newHash);
            if (!equal) throw new InvalidOperationException($"Uncompressed part changed: {oldEntry.FullName}");
            hashes.Add(new { Part = oldEntry.FullName, oldEntry.Length, Sha256 = Convert.ToHexString(oldHash) });
        }
        if (oldZip.Entries.Count != newZip.Entries.Count) throw new InvalidOperationException("Package part count changed.");
        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { Scenario = scenario, BeforeBytes = before.Length, AfterBytes = after.Length, Parts = hashes }));
    }
    return;
}

if (args.Length > 0 && (args[0] == "--profile" || args[0] == "--alloc")) {
    if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("This local profiling harness targets the measured Windows workstation.");
    string scenario = args[1];
    int iterations = int.Parse(args[2]);
    long mask = long.Parse(args[3]);
    Process.GetCurrentProcess().ProcessorAffinity = new IntPtr(mask);
    Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Normal;
    var (instance, method) = Snapshot.Load(scenario, args.Length > 4 ? args[4] : "candidate");
    Action action = Expression.Lambda<Action>(Expression.Block(Expression.Call(Expression.Constant(instance), method), Expression.Empty())).Compile();
    for (int i = 0; i < 32; i++) action();
    Console.WriteLine($"Profile ready: {scenario}; {iterations} iterations; affinity 0x{mask:X}; Normal priority.");
    long allocated = GC.GetAllocatedBytesForCurrentThread();
    for (int i = 0; i < iterations; i++) action();
    allocated = GC.GetAllocatedBytesForCurrentThread() - allocated;
    Console.WriteLine($"Allocated bytes per operation: {allocated / iterations}");
    Console.WriteLine("Profile workload completed with validation.");
    instance.GetType().GetMethod("Cleanup")?.Invoke(instance, null);
    return;
}

var config = DefaultConfig.Instance.AddExporter(JsonExporter.Full)
    .AddJob(Job.Default.WithId("L3-0").WithAffinity(new IntPtr(0xFFFF)))
    .AddJob(Job.Default.WithId("L3-1").WithAffinity(new IntPtr(0xFFFF0000)));
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);

[MemoryDiagnoser]
public class VersionComparison : IDisposable {
    public IEnumerable<string> Scenarios => Environment.GetEnvironmentVariable("OFFICEIMO_COMPARISON_SCENARIOS")?.Split(',')
        ?? ["Csv25K", "Excel25K", "CsvJsonAlways", "CsvQuotesAlways"];
    [ParamsSource(nameof(Scenarios))]
    public string Scenario { get; set; } = "";
    private Func<int> _before = null!;
    private Func<int> _after = null!;
    private int _expectedBefore;
    private int _expectedAfter;
    private object? _beforeInstance;
    private object? _afterInstance;

    [GlobalSetup]
    public void Setup() {
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Normal;
        var before = Snapshot.Load(Scenario, "baseline");
        _beforeInstance = before.Instance;
        var after = Snapshot.Load(Scenario, "candidate");
        _afterInstance = after.Instance;
        static Func<int> Bind((object Instance, MethodInfo Method) target) => target.Method.ReturnType == typeof(int)
            ? (Func<int>)target.Method.CreateDelegate(typeof(Func<int>), target.Instance)
            : Expression.Lambda<Func<int>>(Expression.Block(Expression.Call(Expression.Constant(target.Instance), target.Method), Expression.Constant(0))).Compile();
        _before = Bind(before);
        _after = Bind(after);
        _expectedBefore = _before();
        _expectedAfter = _after();
        if (Scenario.StartsWith("Csv", StringComparison.Ordinal) && _expectedBefore != _expectedAfter) throw new InvalidOperationException("Before/after CSV output lengths differ.");
    }

    [Benchmark(Baseline = true)] public int Before() => Validate(_before(), _expectedBefore);
    [Benchmark] public int After() => Validate(_after(), _expectedAfter);
    private static int Validate(int actual, int expected) => actual == expected ? 0 : throw new InvalidOperationException("Output differs from its validated setup length.");

    [GlobalCleanup]
    public void Dispose() {
        try { _beforeInstance?.GetType().GetMethod("Cleanup")?.Invoke(_beforeInstance, null); }
        finally { _afterInstance?.GetType().GetMethod("Cleanup")?.Invoke(_afterInstance, null); }
    }
}

static class Snapshot {
    public static (object Instance, MethodInfo Method) Load(string scenario, string version) {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PERFORMANCE_AA") == "1") version = "baseline";
        bool csv = scenario.StartsWith("Csv", StringComparison.Ordinal);
        bool read = scenario.EndsWith("Read", StringComparison.Ordinal);
        bool file = scenario.StartsWith("CsvFile", StringComparison.Ordinal);
        bool defaults = scenario.StartsWith("ExcelDefault", StringComparison.Ordinal);
        string kind = csv ? "CSV" : "Excel";
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_PERFORMANCE_SNAPSHOTS") ?? throw new InvalidOperationException("Snapshot root missing.");
        string directory = Path.Combine(root, kind + "-" + version);
        var context = new SnapshotContext(directory);
        string assemblyName = "OfficeIMO." + kind + ".Benchmarks";
        Assembly assembly = context.LoadFromAssemblyPath(Path.Combine(directory, assemblyName + ".dll"));
        string className = file ? "CsvFileWriteBenchmarks" : defaults ? "ExcelDefaultTextWriteBenchmarks" : read ? (csv ? "MarkPflug65KCsvBenchmarks" : "MarkPflug65KXlsxBenchmarks")
            : scenario.EndsWith("25K", StringComparison.Ordinal) ? (csv ? "CsvDataReaderWriteBenchmarks" : "ExcelDataReaderWriteBenchmarks")
            : csv ? "CsvTextWriteBenchmarks" : "ExcelTextWriteBenchmarks";
        Type type = assembly.GetType(assemblyName + "." + className, throwOnError: true)!;
        object instance = Activator.CreateInstance(type)!;
        void Set(string name, object value) {
            PropertyInfo property = type.GetProperty(name)!;
            property.SetValue(instance, property.PropertyType.IsEnum ? Enum.ToObject(property.PropertyType, value) : value);
        }
        string setup = "Setup";
        string method = "OfficeIMO";
        if (file) {
            bool always = scenario.EndsWith("Always", StringComparison.Ordinal);
            Set("Shape", scenario.Substring(7, scenario.Length - 7 - (always ? 6 : 8)));
            Set("QuoteMode", always ? 1 : 0);
        } else if (scenario.EndsWith("25K", StringComparison.Ordinal)) {
            Set("RowCount", 25000);
            if (csv) { Set("Shape", 1); method = "OfficeIMO_WriteDataReader"; }
            setup = "SetupOfficeIMO";
        } else if (!read) {
            Set("TextLength", scenario.Contains("Short", StringComparison.Ordinal) ? 64 : 4096);
            if (csv) Set("QuoteMode", scenario.EndsWith("Always", StringComparison.Ordinal) ? 1 : 0);
            else Set("TextShape", scenario.EndsWith("Markup", StringComparison.Ordinal) ? "Markup" : scenario.EndsWith("Escaped", StringComparison.Ordinal) ? "Escaped" : "Plain");
            if (csv && scenario.Contains("Json", StringComparison.Ordinal)) Set("TextShape", "Json");
            if (csv && scenario.Contains("Quotes", StringComparison.Ordinal)) Set("TextShape", "Quotes");
        }
        type.GetMethod(setup)!.Invoke(instance, null);
        return (instance, type.GetMethod(method)!);
    }
}

sealed class SnapshotContext(string directory) : AssemblyLoadContext {
    protected override Assembly? Load(AssemblyName name) {
        string path = Path.Combine(directory, name.Name + ".dll");
        return File.Exists(path) ? LoadFromAssemblyPath(path) : null;
    }
}
