using System.Diagnostics;
using System.Globalization;
using System.Text;
using DotNet = YamlDotNet.Serialization;
using One = OneYaml;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Benchmarks;

/// <summary>
/// Warm in-process timings of the serializer lanes: the median of repeated rounds, with bytes allocated per operation.
/// A cross-check for BenchmarkDotNet on a busy host, where separate processes see different load.
/// </summary>
internal static class YamlSerializerTimings {
    public static int Run(string[] args) {
        int rounds = 15;
        string? only = null, onlyWorkload = null, onlyOperation = null;
        for (int i = 0; i < args.Length - 1; i++) {
            if (args[i] == "--rounds") rounds = int.Parse(args[i + 1], CultureInfo.InvariantCulture);
            if (args[i] == "--lane") only = args[i + 1];
            if (args[i] == "--workload") onlyWorkload = args[i + 1];
            if (args[i] == "--operation") onlyOperation = args[i + 1];
        }
        var context = new BenchYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase });
        var reflection = new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase };
        BenchSharpYamlContext sharp = BenchSharpYamlContext.Default;
        BenchOneYamlContext one = BenchOneYamlContext.Default;
        DotNet.ISerializer dotNetWriter = new DotNet.SerializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();
        DotNet.IDeserializer dotNetReader = new DotNet.DeserializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();

        Console.WriteLine("| Operation | Workload | Lane | Median | CPU median | Allocated/op |");
        Console.WriteLine("| --- | --- | --- | ---: | ---: | ---: |");
        foreach (string workload in new[] { "SmallConfig", "LargeRecords", "KubernetesNested" }) {
            if (onlyWorkload is not null && workload != onlyWorkload) continue;
            object value = YamlSerializerWorkloads.Create(workload);
            YamlSerializerWorkloads.Validate(value);
            Type type = value.GetType();
            string text = YamlSerializerWorkloads.SerializeOfficeIMO(value);
            byte[] utf8 = Encoding.UTF8.GetBytes(text);
            var lanes = new List<(string Operation, string Lane, Func<object> Action)> {
                ("Serialize", "YamlDotNet", () => dotNetWriter.Serialize(value)),
                ("Serialize", "SharpYaml", () => Sharp.YamlSerializer.Serialize(value, type, sharp)),
                ("Serialize", "VYaml", () => value switch {
                    BenchConfig config => VYaml.Serialization.YamlSerializer.Serialize(config),
                    List<BenchRecord> records => VYaml.Serialization.YamlSerializer.Serialize(records),
                    _ => VYaml.Serialization.YamlSerializer.Serialize((List<BenchDeployment>)value)
                }),
                ("Serialize", "OfficeIMO generated", () => YamlSerializer.Serialize(value, type, context)),
                ("Serialize", "OfficeIMO reflection", () => YamlSerializer.Serialize(value, type, reflection)),
                ("Serialize", "OfficeIMO generated UTF-8", () => {
                    var stream = new MemoryStream();
                    YamlSerializer.Serialize(stream, value, type, context);
                    return stream;
                }),
                ("Serialize", "OfficeIMO generated UTF-8 bytes", () => YamlSerializer.SerializeToUtf8Bytes(value, type, context)),
                ("Serialize", "OneYaml generated", () => One.YamlSerializer.Serialize(value, one.GetTypeInfo(type)!)),
                ("Serialize", "OneYaml reflection", () => One.YamlSerializer.Serialize(value, type)),
                ("Serialize", "OneYaml generated UTF-8", () => value switch {
                    BenchConfig config => One.YamlSerializer.SerializeToUtf8Bytes(config, one.BenchConfig),
                    List<BenchRecord> records => One.YamlSerializer.SerializeToUtf8Bytes(records, one.ListBenchRecord),
                    _ => One.YamlSerializer.SerializeToUtf8Bytes((List<BenchDeployment>)value, one.ListBenchDeployment)
                }),
                ("Deserialize", "YamlDotNet", () => dotNetReader.Deserialize(text, type)!),
                ("Deserialize", "SharpYaml", () => Sharp.YamlSerializer.Deserialize(text, type, sharp)!),
                ("Deserialize", "VYaml", () => value switch {
                    BenchConfig => VYaml.Serialization.YamlSerializer.Deserialize<BenchConfig>(utf8),
                    List<BenchRecord> => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchRecord>>(utf8),
                    _ => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchDeployment>>(utf8)
                }),
                ("Deserialize", "OfficeIMO generated", () => YamlSerializer.Deserialize(text, type, context)!),
                ("Deserialize", "OfficeIMO reflection", () => YamlSerializer.Deserialize(text, type, reflection)!),
                ("Deserialize", "OneYaml generated", () => One.YamlSerializer.Deserialize(text, one.GetTypeInfo(type)!)!),
                ("Deserialize", "OneYaml reflection", () => One.YamlSerializer.Deserialize(text, type)!),
                ("Deserialize", "OneYaml generated UTF-8", () => value switch {
                    BenchConfig => One.YamlSerializer.Deserialize(utf8, one.BenchConfig)!,
                    List<BenchRecord> => One.YamlSerializer.Deserialize(utf8, one.ListBenchRecord)!,
                    _ => One.YamlSerializer.Deserialize(utf8, one.ListBenchDeployment)!
                })
            };
            lanes.RemoveAll(entry => (only is not null && !entry.Lane.Contains(only, StringComparison.OrdinalIgnoreCase))
                || (onlyOperation is not null && entry.Operation != onlyOperation));
            (double[] Samples, double[] Cpu, long Allocated, int Iterations)[] results = Measure(lanes.Select(entry => entry.Action).ToArray(), rounds);
            for (int i = 0; i < lanes.Count; i++) {
                Array.Sort(results[i].Samples);
                Array.Sort(results[i].Cpu);
                Console.WriteLine("| " + lanes[i].Operation + " | " + workload + " | " + lanes[i].Lane + " | " + Format(results[i].Samples[rounds / 2]) + " | "
                    + Format(results[i].Cpu[rounds / 2]) + " | " + results[i].Allocated.ToString("N0", CultureInfo.InvariantCulture) + " B |");
            }
        }
        return 0;
    }

    /// <summary>
    /// Times the lanes round-robin, so load changes on a busy host affect every lane alike: each round runs every lane
    /// for about 50 ms. Besides elapsed time it records the process's processor time, which excludes the time the
    /// thread waits for a core, so a crowded host shifts it far less. Every lane first warms up for a second, so tiered compilation and dynamic PGO finish before
    /// timing, as BenchmarkDotNet's warmup does.
    /// </summary>
    private static (double[] Samples, double[] Cpu, long Allocated, int Iterations)[] Measure(Func<object>[] actions, int rounds) {
        var results = new (double[] Samples, double[] Cpu, long Allocated, int Iterations)[actions.Length];
        Process process = Process.GetCurrentProcess();
        var watch = new Stopwatch();
        for (int lane = 0; lane < actions.Length; lane++) {
            watch.Restart();
            while (watch.ElapsedMilliseconds < 1000) actions[lane]();
            watch.Restart();
            int iterations = 0;
            while (watch.ElapsedMilliseconds < 50) {
                actions[lane]();
                iterations++;
            }
            results[lane] = (new double[rounds], new double[rounds], 0, iterations);
        }
        for (int round = 0; round < rounds; round++) {
            for (int lane = 0; lane < actions.Length; lane++) {
                int iterations = results[lane].Iterations;
                long before = GC.GetAllocatedBytesForCurrentThread();
                process.Refresh();
                TimeSpan cpu = process.TotalProcessorTime;
                watch.Restart();
                for (int i = 0; i < iterations; i++) actions[lane]();
                results[lane].Samples[round] = watch.Elapsed.TotalMilliseconds * 1000 / iterations;
                process.Refresh();
                results[lane].Cpu[round] = (process.TotalProcessorTime - cpu).TotalMilliseconds * 1000 / iterations;
                results[lane].Allocated = (GC.GetAllocatedBytesForCurrentThread() - before) / iterations;
            }
        }
        return results;
    }

    private static string Format(double microseconds) => microseconds >= 1000
        ? (microseconds / 1000).ToString("N2", CultureInfo.InvariantCulture) + " ms"
        : microseconds.ToString("N2", CultureInfo.InvariantCulture) + " µs";
}
