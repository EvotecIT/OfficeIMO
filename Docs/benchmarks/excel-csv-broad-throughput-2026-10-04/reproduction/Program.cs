using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Loader;
using System.Text.Json;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

if (args.Length > 0 && args[0] == "--analyze") { TraceSummary.Run(args[1]); return; }

if (args.Length > 0 && args[0] == "--conformance") {
    string directory = Path.Combine(Environment.GetEnvironmentVariable("OFFICEIMO_PERFORMANCE_SNAPSHOTS")!, "candidate-excel");
    var context = new SnapshotContext(directory);
    var assembly = context.LoadFromAssemblyPath(Path.Combine(directory, "OfficeIMO.Excel.Benchmarks.dll"));
    Type type = assembly.GetType("OfficeIMO.Excel.Benchmarks.ExcelNativeBinaryWriteBenchmarks", true)!;
    foreach (string format in new[] { "Xls", "Xlsb" }) {
        foreach (int rows in new[] { 1, 257, 5000 }) {
            object benchmark = Activator.CreateInstance(type)!;
            type.GetProperty("RowCount")!.SetValue(benchmark, rows);
            var property = type.GetProperty("Format")!;
            property.SetValue(benchmark, Enum.Parse(property.PropertyType, format));
            object result = type.GetMethod("SetupComparison", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(benchmark, null)!;
            Console.WriteLine(JsonSerializer.Serialize(new { Format = format, RowCount = rows, Observation = result }));
        }
    }
    return;
}

if (args.Length > 0 && args[0] == "--profile") {
    Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Normal;
    if (OperatingSystem.IsWindows() || OperatingSystem.IsLinux())
        Process.GetCurrentProcess().ProcessorAffinity = new IntPtr(Convert.ToInt64(Environment.GetEnvironmentVariable("OFFICEIMO_COMPARISON_MASK") ?? "65535"));
    using var probe = SnapshotProbe.Load(args[1], args[2]);
    Console.WriteLine("Setup validated; starting profile operations.");
    for (int i = 0; i < int.Parse(args[3]); i++) await probe.Run();
    Console.WriteLine("Profile operations completed.");
    return;
}

if (args.Length > 0 && args[0] == "--validate") {
    foreach (string key in SnapshotProbe.Keys()) {
        using var before = SnapshotProbe.Load(key, "baseline");
        using var after = SnapshotProbe.Load(key, "candidate");
        string? oldValue = (await before.Run())?.ToString();
        string? newValue = (await after.Run())?.ToString();
        if (oldValue != newValue) throw new InvalidDataException($"{key}: output differs: {oldValue} / {newValue}");
        Console.WriteLine($"Validated {key}: {oldValue}");
    }
    return;
}
long mask = Convert.ToInt64(Environment.GetEnvironmentVariable("OFFICEIMO_COMPARISON_MASK") ?? "65535");
var config = DefaultConfig.Instance.AddExporter(JsonExporter.Full)
    .AddJob(Job.Default.WithId($"Affinity-{mask:X}").WithAffinity(new IntPtr(mask)));
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);

[MemoryDiagnoser]
public class SnapshotComparison {
    public IEnumerable<string> Scenarios => SnapshotProbe.Keys();
    [ParamsSource(nameof(Scenarios))] public string Scenario { get; set; } = "";
    private SnapshotProbe _before = null!, _after = null!;
    [GlobalSetup] public void Setup() {
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Normal;
        _before = SnapshotProbe.Load(Scenario, "baseline");
        _after = SnapshotProbe.Load(Scenario, "candidate");
        if (_before.Run().GetAwaiter().GetResult()?.ToString() != _after.Run().GetAwaiter().GetResult()?.ToString())
            throw new InvalidDataException($"{Scenario}: baseline and candidate outputs differ.");
    }
    [Benchmark(Baseline = true)] public Task<object?> Before() => _before.Run();
    [Benchmark] public Task<object?> After() => _after.Run();
    [GlobalCleanup] public void Cleanup() { try { _before?.Dispose(); } finally { _after?.Dispose(); } }
}

public sealed class SnapshotProbe : IDisposable {
    private readonly object _instance;
    private readonly Func<Task<object?>> _run;
    private SnapshotProbe(object instance, Func<Task<object?>> run) { _instance = instance; _run = run; }
    public Task<object?> Run() => _run();
    public object? RunSynchronously() => _run().GetAwaiter().GetResult();
    private static Dictionary<string, JsonElement> Cases() => JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
        File.ReadAllText(Environment.GetEnvironmentVariable("OFFICEIMO_COMPARISON_CASES")!))!;
    public static IEnumerable<string> Keys() => Cases().Keys;
    public static SnapshotProbe Load(string key, string version) {
        JsonElement spec = Cases()[key];
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PERFORMANCE_AA") == "1") version = "baseline";
        string library = spec.GetProperty("Library").GetString()!;
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_PERFORMANCE_SNAPSHOTS")!;
        string directory = Path.Combine(root, version + "-" + library.ToLowerInvariant());
        var context = new SnapshotContext(directory);
        string assemblyName = $"OfficeIMO.{library}.Benchmarks";
        var assembly = context.LoadFromAssemblyPath(Path.Combine(directory, assemblyName + ".dll"));
        Type type = assembly.GetType(assemblyName + "." + spec.GetProperty("Class").GetString(), true)!;
        object instance = Activator.CreateInstance(type)!;
        foreach (var setting in spec.GetProperty("Properties").EnumerateObject()) {
            var property = type.GetProperty(setting.Name)!;
            object value = property.PropertyType.IsEnum ? Enum.Parse(property.PropertyType, setting.Value.GetString()!)
                : JsonSerializer.Deserialize(setting.Value.GetRawText(), property.PropertyType)!;
            property.SetValue(instance, value);
        }
        object? setup = type.GetMethod(spec.GetProperty("Setup").GetString()!)!.Invoke(instance, null);
        if (setup is Task task) task.GetAwaiter().GetResult();
        var method = type.GetMethod(spec.GetProperty("Method").GetString()!)!;
        Func<Task<object?>> run;
        if (method.ReturnType.IsGenericType && method.ReturnType.GetGenericTypeDefinition() == typeof(Task<>)) {
            run = (Func<Task<object?>>)typeof(SnapshotProbe).GetMethod(nameof(BindAsync), BindingFlags.Static | BindingFlags.NonPublic)!
                .MakeGenericMethod(method.ReturnType.GenericTypeArguments[0]).Invoke(null, [instance, method])!;
        } else {
            var sync = Expression.Lambda<Func<object?>>(Expression.Convert(Expression.Call(Expression.Constant(instance), method), typeof(object))).Compile();
            run = () => Task.FromResult(sync());
        }
        return new SnapshotProbe(instance, run);
    }
    private static Func<Task<object?>> BindAsync<T>(object instance, MethodInfo method) {
        var action = (Func<Task<T>>)method.CreateDelegate(typeof(Func<Task<T>>), instance);
        return async () => await action().ConfigureAwait(false);
    }
    public void Dispose() {
        object? result = _instance.GetType().GetMethod("Cleanup")?.Invoke(_instance, null);
        if (result is Task task) task.GetAwaiter().GetResult();
    }
}
sealed class SnapshotContext(string directory) : AssemblyLoadContext {
    protected override Assembly? Load(AssemblyName name) {
        string path = Path.Combine(directory, name.Name + ".dll");
        return File.Exists(path) ? LoadFromAssemblyPath(path) : null;
    }
}
