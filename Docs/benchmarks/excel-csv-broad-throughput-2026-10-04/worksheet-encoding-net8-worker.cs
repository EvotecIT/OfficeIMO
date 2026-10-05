using System.Data.Common;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;
using System.Security.Cryptography;
using System.Text.Json;

if (args.Length != 8) throw new ArgumentException("Expected root, compare directory, PowerForge DLL, case, side, output, affinity and mode.");
string root = Path.GetFullPath(args[0]);
string compareDirectory = Path.GetFullPath(args[1]);
string powerForgePath = Path.GetFullPath(args[2]);
string caseName = args[3];
string side = args[4];
if (side is not ("Before" or "After")) throw new ArgumentException("Unknown side.");
string output = Path.GetFullPath(args[5]);
long mask = long.Parse(args[6], CultureInfo.InvariantCulture);
bool validateOnly = args[7] == "validate";
if (!validateOnly && args[7] != "measure") throw new ArgumentException("Unknown worker mode.");
using JsonDocument fixtures = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "memory-fixtures", caseName + ".json")));
JsonElement fixture = fixtures.RootElement;
using JsonDocument specifications = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "cases.json")));
JsonElement specification = specifications.RootElement.GetProperty(caseName);
string path = Path.Combine(root, "memory-fixtures", caseName + ".xlsx");
string fileHash = Hash(path);
if (fileHash != fixture.GetProperty("FileSha256").GetString()) throw new InvalidDataException("Memory fixture changed.");
using Process process = Process.GetCurrentProcess();
if (OperatingSystem.IsWindows() || OperatingSystem.IsLinux()) process.ProcessorAffinity = new IntPtr(mask);
process.PriorityClass = ProcessPriorityClass.Normal;
string directory = Path.Combine(root, "net8.0", (side == "Before" ? "baseline" : "candidate") + "-excel");
Assembly harness = AssemblyLoadContext.Default.LoadFromAssemblyPath(Path.Combine(compareDirectory, "bin", "Release", "net8.0", "Compare.dll"));
Type contextType = harness.GetType("SnapshotContext", throwOnError: true)!;
var context = (AssemblyLoadContext)contextType.GetConstructor([typeof(string)])!.Invoke([directory]);
Assembly benchmarks = context.LoadFromAssemblyPath(Path.Combine(directory, "OfficeIMO.Excel.Benchmarks.dll"));
Type type = benchmarks.GetType("OfficeIMO.Excel.Benchmarks." + specification.GetProperty("Class").GetString(), throwOnError: true)!;
object instance = Activator.CreateInstance(type)!;
foreach (JsonProperty setting in specification.GetProperty("Properties").EnumerateObject()) {
    PropertyInfo property = type.GetProperty(setting.Name)!;
    property.SetValue(instance, JsonSerializer.Deserialize(setting.Value.GetRawText(), property.PropertyType));
}
const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
type.GetField("_path", PrivateInstance)!.SetValue(instance, path);
foreach (JsonProperty setting in fixture.GetProperty("State").EnumerateObject()) {
    FieldInfo field = type.GetField(setting.Name, PrivateInstance)!;
    field.SetValue(instance, JsonSerializer.Deserialize(setting.Value.GetRawText(), field.FieldType));
}
MethodInfo method = type.GetMethod(specification.GetProperty("Method").GetString()!)!;
string expected = fixture.GetProperty("Observation").GetString()!;
if (validateOnly) {
    object? observation;
    if (type.Name == "ExcelBufferedTypedReadBenchmarks") {
        MethodInfo read = type.GetMethod("ReadAll", PrivateInstance)!;
        Type modeType = Nullable.GetUnderlyingType(read.GetParameters()[1].ParameterType)!;
        object? mode = method.Name == "Read" ? null : Enum.Parse(modeType, method.Name[4..]);
        observation = read.Invoke(instance, [true, mode]);
    } else if (type.Name == "ExcelNumericXmlReadBenchmarks") {
        observation = type.GetMethod("ReadCore", PrivateInstance)!.Invoke(instance, [true]);
    } else if (type.Name == "ExcelLargeTypedReadBenchmarks") {
        using var reader = (DbDataReader)type.GetMethod("Open", PrivateInstance)!.Invoke(instance, ["OfficeIMO"])!;
        observation = type.GetMethod("Read", PrivateInstance)!.Invoke(instance, [reader, true, false]);
    } else {
        throw new NotSupportedException("No complete-field validation contract for " + type.Name);
    }
    if (Convert.ToString(observation, CultureInfo.InvariantCulture) != expected) throw new InvalidDataException("Validated observation differs.");
    File.WriteAllText(output, JsonSerializer.Serialize(new { Case = caseName, Engine = side, Runtime = RuntimeInformation.FrameworkDescription, CompleteFieldValidation = true, FileSha256 = fileHash, Observed = observation }));
    return;
}
Assembly powerForge = AssemblyLoadContext.Default.LoadFromAssemblyPath(powerForgePath);
MethodInfo start = powerForge.GetType("PowerForge.BenchmarkMemoryProbe", throwOnError: true)!.GetMethod("Start")!;
GC.Collect();
GC.WaitForPendingFinalizers();
long baseline = GC.GetTotalMemory(forceFullCollection: true);
object sample = start.Invoke(null, [5])!;
try {
    long allocated = GC.GetAllocatedBytesForCurrentThread();
    long allThreads = GC.GetTotalAllocatedBytes(precise: true);
    var timer = Stopwatch.StartNew();
    object? observed = method.Invoke(instance, null);
    timer.Stop();
    long totalAllocation = GC.GetTotalAllocatedBytes(precise: true) - allThreads;
    long allocation = GC.GetAllocatedBytesForCurrentThread() - allocated;
    if (Convert.ToString(observed, CultureInfo.InvariantCulture) != expected) throw new InvalidDataException("Complete observation differs.");
    object peak = sample.GetType().GetMethod("Complete")!.Invoke(sample, null)!;
    long retained = GC.GetTotalMemory(forceFullCollection: true) - baseline;
    File.WriteAllText(output, JsonSerializer.Serialize(new {
        Case = caseName, Engine = side, Pid = Environment.ProcessId,
        Runtime = RuntimeInformation.FrameworkDescription, Host = Environment.MachineName,
        Mask = OperatingSystem.IsWindows() || OperatingSystem.IsLinux() ? mask.ToString(CultureInfo.InvariantCulture) : "OS scheduled",
        FileSha256 = fileHash, ExcelDllSha256 = Hash(Path.Combine(directory, "OfficeIMO.Excel.dll")), PowerForgeSha256 = Hash(powerForgePath),
        ExpectedObservation = expected, Observed = observed, ReadMs = timer.Elapsed.TotalMilliseconds,
        CallingThreadAllocatedBytes = allocation, AllThreadAllocatedBytesIncludingSampler = totalAllocation,
        AfterReturnManagedIncreaseBytes = retained, Peak = peak,
        Contract = "First complete read in a fresh actual .NET8 worker. Separate workers validate every projected field before measurement.",
        Limits = "JIT/type initialization and reflection invocation included; caller allocation excludes sampler/prefetch; all-thread allocation includes sampler; sampled peaks are lower bounds; producer warms OS caches."
    }, new JsonSerializerOptions { WriteIndented = true }));
} finally {
    ((IDisposable)sample).Dispose();
}

static string Hash(string path) {
    using FileStream stream = File.OpenRead(path);
    return Convert.ToHexString(SHA256.HashData(stream));
}
