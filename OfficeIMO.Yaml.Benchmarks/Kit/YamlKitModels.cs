using System.Text.Json.Serialization;
using VYaml.Annotations;
using One = OneYaml;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Benchmarks.Kit;

// Typed models for the oneyaml-kit typed-load workloads: the records shape and the actions/checkout workflow, with the
// fields and shapes the kit's expected values hold. Keys are lower camel case for every library; the three keys that are
// not (runs-on, pull_request, if) carry each library's own name attribute. OneYaml and SharpYaml read the System.Text.Json
// attribute. A field that holds either of two shapes (a job's container, a step's with values) is an object.

[YamlObject]
public partial class KitRecord {
    public long Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public bool Active { get; set; }
    public double Score { get; set; }
    public List<string> Tags { get; set; } = new();
    public KitAddress Address { get; set; } = new();
    public string? Note { get; set; }
}

[YamlObject]
public partial class KitAddress {
    public string City { get; set; } = string.Empty;
    public string Zip { get; set; } = string.Empty;
    public KitGeo? Geo { get; set; }
}

[YamlObject]
public partial class KitGeo {
    public double Lat { get; set; }
    public double Lng { get; set; }
}

[YamlObject]
public partial class KitWorkflow {
    public string Name { get; set; } = string.Empty;
    public KitTriggers On { get; set; } = new();
    public Dictionary<string, KitJob> Jobs { get; set; } = new();
}

[YamlObject]
public partial class KitTriggers {
    [YamlMember("pull_request"), OfficePropertyName("pull_request"), JsonPropertyName("pull_request")]
    public string? PullRequest { get; set; }
    public KitPush Push { get; set; } = new();
}

[YamlObject]
public partial class KitPush {
    public List<string> Branches { get; set; } = new();
}

[YamlObject]
public partial class KitJob {
    [YamlMember("runs-on"), OfficePropertyName("runs-on"), JsonPropertyName("runs-on")]
    public string RunsOn { get; set; } = string.Empty;
    public KitStrategy? Strategy { get; set; }
    /// <summary>An image name or a mapping with <c>image</c> and <c>options</c>.</summary>
    public object? Container { get; set; }
    public Dictionary<string, KitService>? Services { get; set; }
    public Dictionary<string, string>? Env { get; set; }
    public List<KitStep> Steps { get; set; } = new();
}

[YamlObject]
public partial class KitStrategy {
    public KitMatrix Matrix { get; set; } = new();
}

[YamlObject]
public partial class KitMatrix {
    [YamlMember("runs-on"), OfficePropertyName("runs-on"), JsonPropertyName("runs-on")]
    public List<string> RunsOn { get; set; } = new();
}

[YamlObject]
public partial class KitService {
    public string Image { get; set; } = string.Empty;
    public List<string> Ports { get; set; } = new();
}

[YamlObject]
public partial class KitStep {
    public string? Id { get; set; }
    public string? Name { get; set; }
    [YamlMember("if"), OfficePropertyName("if"), JsonPropertyName("if")]
    public string? Condition { get; set; }
    public string? Uses { get; set; }
    public string? Run { get; set; }
    public string? Shell { get; set; }
    /// <summary>Each value is a string or a boolean.</summary>
    public Dictionary<string, object?>? With { get; set; }
}

[YamlSerializable(typeof(KitWorkflow))]
[YamlSerializable(typeof(List<KitRecord>))]
[YamlSerializable(typeof(object))]
public partial class KitYamlContext : YamlSerializerContext {
}

[One.YamlSerializable(typeof(KitWorkflow))]
[One.YamlSerializable(typeof(List<KitRecord>))]
[One.YamlSerializable(typeof(object))]
public partial class KitOneYamlContext : One.YamlSerializerContext {
}

[Sharp.Serialization.YamlSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[Sharp.Serialization.YamlSerializable(typeof(KitWorkflow))]
[Sharp.Serialization.YamlSerializable(typeof(List<KitRecord>))]
public partial class KitSharpYamlContext : Sharp.Serialization.YamlSerializerContext {
}
