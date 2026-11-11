using System.Text;
using BenchmarkDotNet.Attributes;
using VYaml.Annotations;
using DotNet = YamlDotNet.Serialization;
using One = OneYaml;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Benchmarks;

// One model set for every library: VYaml's generator needs partial [YamlObject] types, and every lane uses camelCase
// keys (VYaml's and OneYaml's default). OfficeIMO.Yaml and SharpYaml use their generated contexts, YamlDotNet its
// reflection serializer, VYaml its generated formatters, and OneYaml runs with its generated context and with reflection.

[YamlObject]
public partial class BenchConfig {
    public string Name { get; set; } = string.Empty;
    public string Environment { get; set; } = string.Empty;
    public int Port { get; set; }
    public bool Enabled { get; set; }
    public double Ratio { get; set; }
    public int MaxConnections { get; set; }
    public List<string> Hosts { get; set; } = new List<string>();
    public Dictionary<string, string> Labels { get; set; } = new Dictionary<string, string>();
    public BenchLogging Logging { get; set; } = new BenchLogging();
}

[YamlObject]
public partial class BenchLogging {
    public string Level { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
    public bool Console { get; set; }
}

[YamlObject]
public partial class BenchRecord {
    public int Id { get; set; }
    public string Customer { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public double Discount { get; set; }
    public int Quantity { get; set; }
    public bool Paid { get; set; }
    public string Status { get; set; } = string.Empty;
}

[YamlObject]
public partial class BenchDeployment {
    public string ApiVersion { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public BenchMetadata Metadata { get; set; } = new BenchMetadata();
    public BenchDeploymentSpec Spec { get; set; } = new BenchDeploymentSpec();
}

[YamlObject]
public partial class BenchMetadata {
    public string Name { get; set; } = string.Empty;
    public string Namespace { get; set; } = string.Empty;
    public Dictionary<string, string> Labels { get; set; } = new Dictionary<string, string>();
}

[YamlObject]
public partial class BenchDeploymentSpec {
    public int Replicas { get; set; }
    public BenchSelector Selector { get; set; } = new BenchSelector();
    public BenchPodTemplate Template { get; set; } = new BenchPodTemplate();
}

[YamlObject]
public partial class BenchSelector {
    public Dictionary<string, string> MatchLabels { get; set; } = new Dictionary<string, string>();
}

[YamlObject]
public partial class BenchPodTemplate {
    public BenchMetadata Metadata { get; set; } = new BenchMetadata();
    public BenchPodSpec Spec { get; set; } = new BenchPodSpec();
}

[YamlObject]
public partial class BenchPodSpec {
    public List<BenchContainer> Containers { get; set; } = new List<BenchContainer>();
}

[YamlObject]
public partial class BenchContainer {
    public string Name { get; set; } = string.Empty;
    public string Image { get; set; } = string.Empty;
    public List<string> Args { get; set; } = new List<string>();
    public List<BenchPort> Ports { get; set; } = new List<BenchPort>();
    public List<BenchEnv> Env { get; set; } = new List<BenchEnv>();
    public BenchResources Resources { get; set; } = new BenchResources();
}

[YamlObject]
public partial class BenchPort {
    public int ContainerPort { get; set; }
    public string Protocol { get; set; } = string.Empty;
}

[YamlObject]
public partial class BenchEnv {
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

[YamlObject]
public partial class BenchResources {
    public Dictionary<string, string> Requests { get; set; } = new Dictionary<string, string>();
    public Dictionary<string, string> Limits { get; set; } = new Dictionary<string, string>();
}

[YamlSerializable(typeof(BenchConfig))]
[YamlSerializable(typeof(List<BenchRecord>))]
[YamlSerializable(typeof(List<BenchDeployment>))]
public partial class BenchYamlContext : YamlSerializerContext {
}

[Sharp.Serialization.YamlSourceGenerationOptions(PropertyNamingPolicy = System.Text.Json.Serialization.JsonKnownNamingPolicy.CamelCase)]
[Sharp.Serialization.YamlSerializable(typeof(BenchConfig))]
[Sharp.Serialization.YamlSerializable(typeof(List<BenchRecord>))]
[Sharp.Serialization.YamlSerializable(typeof(List<BenchDeployment>))]
public partial class BenchSharpYamlContext : Sharp.Serialization.YamlSerializerContext {
}

// OneYaml's generator reads its own attribute; it runs beside OfficeIMO.Yaml's and SharpYaml's generators.
[One.YamlSerializable(typeof(BenchConfig))]
[One.YamlSerializable(typeof(List<BenchRecord>))]
[One.YamlSerializable(typeof(List<BenchDeployment>))]
public partial class BenchOneYamlContext : One.YamlSerializerContext {
}

/// <summary>The three serializer workloads, each with an operation per library.</summary>
public abstract class YamlSerializerBenchmarkBase {
    protected static readonly BenchYamlContext Context = new BenchYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase });
    protected static readonly YamlSerializerOptions Reflection = new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase };
    protected static readonly BenchSharpYamlContext SharpContext = BenchSharpYamlContext.Default;
    protected static readonly BenchOneYamlContext OneContext = BenchOneYamlContext.Default;
    protected static readonly DotNet.ISerializer DotNetSerializer = new DotNet.SerializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();
    protected static readonly DotNet.IDeserializer DotNetDeserializer = new DotNet.DeserializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();

    /// <summary>The workload: one configuration POCO, 10,000 flat records, or 200 Kubernetes-like deployments.</summary>
    [Params("SmallConfig", "LargeRecords", "KubernetesNested")]
    public string Workload { get; set; } = "SmallConfig";

    protected object Value = null!;
    protected string Text = string.Empty;
    protected byte[] Utf8 = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        Value = YamlSerializerWorkloads.Create(Workload);
        Text = YamlSerializerWorkloads.SerializeOfficeIMO(Value);
        Utf8 = Encoding.UTF8.GetBytes(Text);
        YamlSerializerWorkloads.Validate(Value);
    }
}

[MemoryDiagnoser]
public class YamlSerializeBenchmarks : YamlSerializerBenchmarkBase {
    [Benchmark(Baseline = true)]
    public int YamlDotNet() => DotNetSerializer.Serialize(Value).Length;

    [Benchmark]
    public int SharpYamlGenerated() => Workload switch {
        "SmallConfig" => Sharp.YamlSerializer.Serialize((BenchConfig)Value, SharpContext).Length,
        "LargeRecords" => Sharp.YamlSerializer.Serialize((List<BenchRecord>)Value, SharpContext).Length,
        _ => Sharp.YamlSerializer.Serialize((List<BenchDeployment>)Value, SharpContext).Length
    };

    [Benchmark]
    public int VYamlGenerated() => Workload switch {
        "SmallConfig" => VYaml.Serialization.YamlSerializer.Serialize((BenchConfig)Value).Length,
        "LargeRecords" => VYaml.Serialization.YamlSerializer.Serialize((List<BenchRecord>)Value).Length,
        _ => VYaml.Serialization.YamlSerializer.Serialize((List<BenchDeployment>)Value).Length
    };

    [Benchmark]
    public int OneYamlGenerated() => Workload switch {
        "SmallConfig" => One.YamlSerializer.Serialize((BenchConfig)Value, OneContext.BenchConfig).Length,
        "LargeRecords" => One.YamlSerializer.Serialize((List<BenchRecord>)Value, OneContext.ListBenchRecord).Length,
        _ => One.YamlSerializer.Serialize((List<BenchDeployment>)Value, OneContext.ListBenchDeployment).Length
    };

    [Benchmark]
    public int OneYamlReflection() => One.YamlSerializer.Serialize(Value, Value.GetType()).Length;

    /// <summary>UTF-8 output as a new array, OneYaml's native form.</summary>
    [Benchmark]
    public int OneYamlGeneratedUtf8() => Workload switch {
        "SmallConfig" => One.YamlSerializer.SerializeToUtf8Bytes((BenchConfig)Value, OneContext.BenchConfig).Length,
        "LargeRecords" => One.YamlSerializer.SerializeToUtf8Bytes((List<BenchRecord>)Value, OneContext.ListBenchRecord).Length,
        _ => One.YamlSerializer.SerializeToUtf8Bytes((List<BenchDeployment>)Value, OneContext.ListBenchDeployment).Length
    };

    [Benchmark]
    public int OfficeIMOGenerated() => YamlSerializerWorkloads.SerializeOfficeIMO(Value).Length;

    [Benchmark]
    public int OfficeIMOReflection() => YamlSerializer.Serialize(Value, Value.GetType(), Reflection).Length;

    /// <summary>UTF-8 output, the form VYaml returns, written to a new stream as VYaml returns new memory.</summary>
    [Benchmark]
    public long OfficeIMOGeneratedUtf8() {
        var stream = new MemoryStream();
        YamlSerializer.Serialize(stream, Value, Value.GetType(), Context);
        return stream.Length;
    }
}

[MemoryDiagnoser]
public class YamlDeserializeBenchmarks : YamlSerializerBenchmarkBase {
    [Benchmark(Baseline = true)]
    public object YamlDotNet() => DotNetDeserializer.Deserialize(Text, Value.GetType())!;

    [Benchmark]
    public object SharpYamlGenerated() => Workload switch {
        "SmallConfig" => Sharp.YamlSerializer.Deserialize<BenchConfig>(Text, SharpContext)!,
        "LargeRecords" => Sharp.YamlSerializer.Deserialize<List<BenchRecord>>(Text, SharpContext)!,
        _ => Sharp.YamlSerializer.Deserialize<List<BenchDeployment>>(Text, SharpContext)!
    };

    [Benchmark]
    public object VYamlGenerated() => Workload switch {
        "SmallConfig" => VYaml.Serialization.YamlSerializer.Deserialize<BenchConfig>(Utf8),
        "LargeRecords" => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchRecord>>(Utf8),
        _ => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchDeployment>>(Utf8)
    };

    [Benchmark]
    public object OneYamlGenerated() => Workload switch {
        "SmallConfig" => One.YamlSerializer.Deserialize(Text, OneContext.BenchConfig)!,
        "LargeRecords" => One.YamlSerializer.Deserialize(Text, OneContext.ListBenchRecord)!,
        _ => One.YamlSerializer.Deserialize(Text, OneContext.ListBenchDeployment)!
    };

    [Benchmark]
    public object OneYamlReflection() => One.YamlSerializer.Deserialize(Text, Value.GetType())!;

    /// <summary>UTF-8 input, OneYaml's native form, encoded once during setup as for VYaml.</summary>
    [Benchmark]
    public object OneYamlGeneratedUtf8() => Workload switch {
        "SmallConfig" => One.YamlSerializer.Deserialize(Utf8, OneContext.BenchConfig)!,
        "LargeRecords" => One.YamlSerializer.Deserialize(Utf8, OneContext.ListBenchRecord)!,
        _ => One.YamlSerializer.Deserialize(Utf8, OneContext.ListBenchDeployment)!
    };

    [Benchmark]
    public object OfficeIMOGenerated() => YamlSerializer.Deserialize(Text, Value.GetType(), Context)!;

    [Benchmark]
    public object OfficeIMOReflection() => YamlSerializer.Deserialize(Text, Value.GetType(), Reflection)!;
}

/// <summary>Builds the workloads and checks every lane's output before timing.</summary>
internal static class YamlSerializerWorkloads {
    private static readonly BenchYamlContext Context = new BenchYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase });

    public static object Create(string workload) => workload switch {
        "SmallConfig" => Config(),
        "LargeRecords" => Enumerable.Range(1, 10_000).Select(Record).ToList(),
        "KubernetesNested" => Enumerable.Range(1, 200).Select(Deployment).ToList(),
        _ => throw new ArgumentOutOfRangeException(nameof(workload), workload, "Unknown workload.")
    };

    public static string SerializeOfficeIMO(object value) => YamlSerializer.Serialize(value, value.GetType(), Context);

    /// <summary>
    /// Every lane's output must read back with OfficeIMO.Yaml to the same objects, and every lane must read OfficeIMO's
    /// output to the same objects; objects are compared through OfficeIMO's deterministic output.
    /// </summary>
    public static void Validate(object value) {
        Type type = value.GetType();
        string expected = SerializeOfficeIMO(value);
        var sharp = BenchSharpYamlContext.Default;
        DotNet.ISerializer dotNetWriter = new DotNet.SerializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();
        DotNet.IDeserializer dotNetReader = new DotNet.DeserializerBuilder().WithNamingConvention(DotNet.NamingConventions.CamelCaseNamingConvention.Instance).Build();
        var reflection = new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase };
        BenchOneYamlContext one = BenchOneYamlContext.Default;
        var outputs = new Dictionary<string, string> {
            ["OfficeIMO reflection"] = YamlSerializer.Serialize(value, type, reflection),
            ["OneYaml generated"] = One.YamlSerializer.Serialize(value, one.GetTypeInfo(type)!),
            ["OneYaml reflection"] = One.YamlSerializer.Serialize(value, type),
            ["OneYaml generated UTF-8"] = Encoding.UTF8.GetString(value switch {
                BenchConfig config => One.YamlSerializer.SerializeToUtf8Bytes(config, one.BenchConfig),
                List<BenchRecord> records => One.YamlSerializer.SerializeToUtf8Bytes(records, one.ListBenchRecord),
                _ => One.YamlSerializer.SerializeToUtf8Bytes((List<BenchDeployment>)value, one.ListBenchDeployment)
            }),
            ["YamlDotNet"] = dotNetWriter.Serialize(value),
            ["SharpYaml"] = Sharp.YamlSerializer.Serialize(value, type, sharp),
            ["VYaml"] = value switch {
                BenchConfig config => VYaml.Serialization.YamlSerializer.SerializeToString(config),
                List<BenchRecord> records => VYaml.Serialization.YamlSerializer.SerializeToString(records),
                _ => VYaml.Serialization.YamlSerializer.SerializeToString((List<BenchDeployment>)value)
            }
        };
        foreach (KeyValuePair<string, string> output in outputs) {
            Check(output.Key + " output", expected, SerializeOfficeIMO(YamlSerializer.Deserialize(output.Value, type, Context)!));
        }
        Check("YamlDotNet read", expected, SerializeOfficeIMO(dotNetReader.Deserialize(expected, type)!));
        Check("SharpYaml read", expected, SerializeOfficeIMO(Sharp.YamlSerializer.Deserialize(expected, type, sharp)!));
        byte[] utf8 = Encoding.UTF8.GetBytes(expected);
        object vyaml = value switch {
            BenchConfig => VYaml.Serialization.YamlSerializer.Deserialize<BenchConfig>(utf8),
            List<BenchRecord> => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchRecord>>(utf8),
            _ => VYaml.Serialization.YamlSerializer.Deserialize<List<BenchDeployment>>(utf8)
        };
        Check("VYaml read", expected, SerializeOfficeIMO(vyaml));
        Check("OneYaml generated read", expected, SerializeOfficeIMO(One.YamlSerializer.Deserialize(expected, one.GetTypeInfo(type)!)!));
        Check("OneYaml reflection read", expected, SerializeOfficeIMO(One.YamlSerializer.Deserialize(expected, type)!));
        object oneUtf8 = value switch {
            BenchConfig => One.YamlSerializer.Deserialize(utf8, one.BenchConfig)!,
            List<BenchRecord> => One.YamlSerializer.Deserialize(utf8, one.ListBenchRecord)!,
            _ => One.YamlSerializer.Deserialize(utf8, one.ListBenchDeployment)!
        };
        Check("OneYaml generated UTF-8 read", expected, SerializeOfficeIMO(oneUtf8));
    }

    private static void Check(string lane, string expected, string actual) {
        if (expected != actual) throw new InvalidOperationException(lane + " does not match the workload's data.");
    }

    private static BenchConfig Config() => new BenchConfig {
        Name = "orders-api",
        Environment = "production",
        Port = 8443,
        Enabled = true,
        Ratio = 0.75,
        MaxConnections = 512,
        Hosts = { "api.example.com", "api-eu.example.com", "10.0.0.12" },
        Labels = { ["team"] = "payments", ["tier"] = "backend", ["region"] = "eu-west-1" },
        Logging = new BenchLogging { Level = "Information", Format = "json", Console = true }
    };

    private static BenchRecord Record(int id) => new BenchRecord {
        Id = id,
        Customer = "Customer " + (id % 977).ToString(System.Globalization.CultureInfo.InvariantCulture),
        Email = "customer" + id.ToString(System.Globalization.CultureInfo.InvariantCulture) + "@example.com",
        Amount = id * 1.25m,
        Discount = (id % 10) / 20.0,
        Quantity = id % 13 + 1,
        Paid = id % 3 != 0,
        Status = (id % 4) switch { 0 => "pending", 1 => "shipped", 2 => "delivered", _ => "returned" }
    };

    private static BenchDeployment Deployment(int index) {
        string name = "service-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var labels = new Dictionary<string, string> { ["app"] = name, ["tier"] = index % 2 == 0 ? "frontend" : "backend" };
        var deployment = new BenchDeployment {
            ApiVersion = "apps/v1",
            Kind = "Deployment",
            Metadata = new BenchMetadata { Name = name, Namespace = "production", Labels = new Dictionary<string, string>(labels) },
            Spec = new BenchDeploymentSpec {
                Replicas = index % 5 + 1,
                Selector = new BenchSelector { MatchLabels = new Dictionary<string, string>(labels) },
                Template = new BenchPodTemplate { Metadata = new BenchMetadata { Name = name, Namespace = "production", Labels = labels } }
            }
        };
        for (int c = 0; c < 2; c++) {
            deployment.Spec.Template.Spec.Containers.Add(new BenchContainer {
                Name = c == 0 ? "app" : "sidecar",
                Image = "registry.example.com/" + name + ":1." + (index % 30).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".0",
                Args = { "--port", "8080", "--log-level", "info" },
                Ports = { new BenchPort { ContainerPort = 8080 + c, Protocol = "TCP" } },
                Env = { new BenchEnv { Name = "MODE", Value = "production" }, new BenchEnv { Name = "REPLICA", Value = index.ToString(System.Globalization.CultureInfo.InvariantCulture) } },
                Resources = new BenchResources {
                    Requests = { ["cpu"] = "100m", ["memory"] = "128Mi" },
                    Limits = { ["cpu"] = "500m", ["memory"] = "512Mi" }
                }
            });
        }
        return deployment;
    }
}
