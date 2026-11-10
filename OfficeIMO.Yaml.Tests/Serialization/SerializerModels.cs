using System.Numerics;
using OfficeIMO.Yaml;

namespace OfficeIMO.Yaml.Tests.Serialization;

// The representative model set: each shape the serializer supports appears at least once, and the same types are read
// and written through generated and reflection metadata.

public enum LogLevel { Trace, Debug, Information, Warning, Error }

[Flags]
public enum Permissions { None = 0, Read = 1, Write = 2, Execute = 4, ReadWrite = Read | Write }

public sealed class AppConfig {
    public string Name { get; set; } = string.Empty;
    public int Port { get; set; }
    public bool Enabled { get; set; }
    public LogLevel Level { get; set; }
    public Permissions Access { get; set; }
    public double Ratio { get; set; }
    public decimal Price { get; set; }
    public TimeSpan Timeout { get; set; }
    public Uri? Endpoint { get; set; }
    public int? Retries { get; set; }
    public string? Comment { get; set; }
    public List<string> Tags { get; set; } = new List<string>();
    public Dictionary<string, string> Labels { get; set; } = new Dictionary<string, string>();
    public string[] Hosts { get; set; } = Array.Empty<string>();

    [YamlPropertyName("feature-flags")]
    public Dictionary<string, bool> Features { get; set; } = new Dictionary<string, bool>();

    [YamlIgnore]
    public string Secret { get; set; } = "never written";
}

public sealed class ConfigSet {
    public AppConfig? Base { get; set; }
    public List<AppConfig> Configs { get; set; } = new List<AppConfig>();
}

public sealed record Person(string Name, int Age, string? Email = null);

public sealed class Account {
    public required string Id { get; init; }
    public string Owner { get; init; } = "unknown";
    [YamlRequired]
    public decimal Balance { get; set; }
    public DateTimeOffset Opened { get; set; }
}

public struct Point {
    public int X { get; set; }
    public int Y { get; set; }
}

public readonly struct Size {
    [YamlConstructor]
    public Size(int width, int height) {
        Width = width;
        Height = height;
    }

    public int Width { get; }
    public int Height { get; }
}

public sealed class Scalars {
    public string Text { get; set; } = string.Empty;
    public char Letter { get; set; }
    public byte Byte { get; set; }
    public sbyte SByte { get; set; }
    public short Short { get; set; }
    public ushort UShort { get; set; }
    public uint UInt { get; set; }
    public long Long { get; set; }
    public ulong ULong { get; set; }
    public float Single { get; set; }
    public double Double { get; set; }
    public decimal Decimal { get; set; }
    public BigInteger Big { get; set; }
    public DateTime When { get; set; }
    public DateTimeOffset At { get; set; }
    public TimeSpan Duration { get; set; }
    public Guid Id { get; set; }
    public byte[] Payload { get; set; } = Array.Empty<byte>();
    public Point Origin { get; set; }
    public Point? Maybe { get; set; }
    public Size Box { get; set; }
#if NET6_0_OR_GREATER
    public DateOnly Day { get; set; }
    public TimeOnly Clock { get; set; }
#endif
}

public sealed class Collections {
    public int[] Numbers { get; set; } = Array.Empty<int>();
    public IReadOnlyList<string> ReadOnly { get; set; } = new List<string>();
    public IEnumerable<long> Sequence { get; set; } = new List<long>();
    public HashSet<string> Unique { get; set; } = new HashSet<string>();
    public ISet<int> Set { get; set; } = new HashSet<int>();
    public LinkedList<string> Linked { get; set; } = new LinkedList<string>();
    public SortedDictionary<string, int> Sorted { get; set; } = new SortedDictionary<string, int>();
    public IReadOnlyDictionary<int, string> ByNumber { get; set; } = new Dictionary<int, string>();
    public Dictionary<LogLevel, List<string>> ByLevel { get; set; } = new Dictionary<LogLevel, List<string>>();
    public List<Dictionary<string, Point>> Nested { get; set; } = new List<Dictionary<string, Point>>();
    public object? Anything { get; set; }
}

[YamlDerivedType(typeof(Cat), "cat")]
[YamlDerivedType(typeof(Dog), "dog")]
public abstract class Animal {
    public string Name { get; set; } = string.Empty;
}

public sealed class Cat : Animal {
    public int Lives { get; set; }
}

public sealed class Dog : Animal {
    public bool GoodBoy { get; set; }
}

public sealed class Zoo {
    public List<Animal> Animals { get; set; } = new List<Animal>();
    public Animal? Favorite { get; set; }
}

public sealed class TreeNode {
    public string Name { get; set; } = string.Empty;
    public TreeNode? Parent { get; set; }
    public List<TreeNode> Children { get; set; } = new List<TreeNode>();
}

// A Kubernetes-like nested model, with the object kind as a property discriminator ("kind" is the second key in real files).
[YamlPolymorphic(DiscriminatorPropertyName = "kind")]
[YamlDerivedType(typeof(KubeDeployment), "Deployment")]
[YamlDerivedType(typeof(KubeService), "Service")]
public abstract class KubeObject {
    public string ApiVersion { get; set; } = string.Empty;
    public KubeMetadata Metadata { get; set; } = new KubeMetadata();
}

public sealed class KubeMetadata {
    public string Name { get; set; } = string.Empty;
    public string? Namespace { get; set; }
    public Dictionary<string, string>? Labels { get; set; }
    public Dictionary<string, string>? Annotations { get; set; }
}

public sealed class KubeDeployment : KubeObject {
    public DeploymentSpec Spec { get; set; } = new DeploymentSpec();
}

public sealed class DeploymentSpec {
    public int? Replicas { get; set; }
    public LabelSelector? Selector { get; set; }
    public PodTemplate Template { get; set; } = new PodTemplate();
}

public sealed class LabelSelector {
    public Dictionary<string, string> MatchLabels { get; set; } = new Dictionary<string, string>();
}

public sealed class PodTemplate {
    public KubeMetadata Metadata { get; set; } = new KubeMetadata();
    public PodSpec Spec { get; set; } = new PodSpec();
}

public sealed class PodSpec {
    public List<Container> Containers { get; set; } = new List<Container>();
    public string? ServiceAccountName { get; set; }
    public Dictionary<string, string>? NodeSelector { get; set; }
}

public sealed class Container {
    public string Name { get; set; } = string.Empty;
    public string Image { get; set; } = string.Empty;
    public string? ImagePullPolicy { get; set; }
    public List<string>? Args { get; set; }
    public List<PortSpec>? Ports { get; set; }
    public List<EnvVar>? Env { get; set; }
    public ResourceRequirements? Resources { get; set; }
}

public sealed class PortSpec {
    public int ContainerPort { get; set; }
    public string? Name { get; set; }
    public string? Protocol { get; set; }
}

public sealed class EnvVar {
    public string Name { get; set; } = string.Empty;
    public string? Value { get; set; }
}

public sealed class ResourceRequirements {
    public Dictionary<string, string>? Requests { get; set; }
    public Dictionary<string, string>? Limits { get; set; }
}

public sealed class KubeService : KubeObject {
    public ServiceSpec Spec { get; set; } = new ServiceSpec();
}

public sealed class ServiceSpec {
    public string? Type { get; set; }
    public List<ServicePort> Ports { get; set; } = new List<ServicePort>();
    public Dictionary<string, string>? Selector { get; set; }
}

public sealed class ServicePort {
    public int Port { get; set; }
    public int? TargetPort { get; set; }
    public string? Protocol { get; set; }
}

// The Compose file shape: services by name, with the environment and volumes Airflow's file shares through anchors.
public sealed class ComposeFile {
    public Dictionary<string, ComposeService> Services { get; set; } = new Dictionary<string, ComposeService>();
    public Dictionary<string, object?>? Volumes { get; set; }
}

public sealed class ComposeService {
    public string? Image { get; set; }
    public string? Command { get; set; }
    public Dictionary<string, string>? Environment { get; set; }
    public List<string>? Volumes { get; set; }
    public List<string>? Ports { get; set; }
    public string? Restart { get; set; }
}

[YamlSerializable(typeof(AppConfig))]
[YamlSerializable(typeof(ConfigSet))]
[YamlSerializable(typeof(Person))]
[YamlSerializable(typeof(List<Person>))]
[YamlSerializable(typeof(Account))]
[YamlSerializable(typeof(Scalars))]
[YamlSerializable(typeof(Collections))]
[YamlSerializable(typeof(Zoo))]
[YamlSerializable(typeof(TreeNode))]
[YamlSerializable(typeof(KubeObject))]
[YamlSerializable(typeof(ComposeFile))]
[YamlSerializable(typeof(Dictionary<string, object>))]
public partial class TestYamlContext : YamlSerializerContext {
}
