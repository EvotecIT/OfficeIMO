using OfficeIMO.Yaml;

// Serialization through generated metadata only: no reflection, so trimming and NativeAOT keep every member it needs.
var context = new SmokeYamlContext(new YamlSerializerOptions {
    PropertyNamingPolicy = YamlNamingPolicy.CamelCase,
    DefaultIgnoreCondition = YamlIgnoreCondition.WhenWritingNull
});

var config = new ServiceConfig {
    Name = "web",
    Replicas = 3,
    Level = Level.Warning,
    Access = Access.Read | Access.Write,
    Timeout = TimeSpan.FromSeconds(30),
    Owner = new Owner("Ada", "ada@example.com"),
    Labels = { ["app"] = "web", ["tier"] = "yes" },
    Ports = { new Port { Number = 80 }, new Port { Number = 443, Name = "https" } },
    Probes = { new HttpProbe { Path = "/health" }, new TcpProbe { Port = 5432 } },
    Created = new DateTimeOffset(2026, 11, 10, 9, 0, 0, TimeSpan.Zero),
    Id = "svc-1"
};

const string expected = "name: web\nreplicas: 3\nlevel: Warning\naccess: Read, Write\ntimeout: 00:00:30\n"
    + "owner:\n  name: Ada\n  email: ada@example.com\nlabels:\n  app: web\n  tier: 'yes'\n"
    + "ports:\n- number: 80\n- number: 443\n  name: https\nprobes:\n- !http\n  path: /health\n- !tcp\n  port: 5432\n"
    + "created: 2026-11-10T09:00:00+00:00\nid: svc-1\n";
string yaml = YamlSerializer.Serialize(config, context.ServiceConfig);
Require(yaml == expected, "The generated serializer wrote unexpected YAML:\n" + yaml);

ServiceConfig copy = YamlSerializer.Deserialize(yaml, context.ServiceConfig)
    ?? throw new InvalidOperationException("The generated serializer read nothing.");
Require(copy.Owner == config.Owner && copy.Access == config.Access && copy.Labels["tier"] == "yes"
    && copy.Probes[0] is HttpProbe { Path: "/health" } && copy.Probes[1] is TcpProbe { Port: 5432 } && copy.Id == "svc-1",
    "The generated serializer lost a value.");
Require(YamlSerializer.Serialize(copy, context.ServiceConfig) == yaml, "The generated serializer did not round-trip.");

// Aliases, merge keys and untyped values.
ServiceConfig merged = YamlSerializer.Deserialize("defaults: &d {name: base, replicas: 2, id: x}\nname: override\n<<: *d\n", context.ServiceConfig)!;
Require(merged.Name == "override" && merged.Replicas == 2, "Merge keys were not applied.");
var untyped = (Dictionary<string, object?>)YamlSerializer.Deserialize("a: 1\nb: [true, 2.5, text]\n", typeof(object), context)!;
Require(untyped["a"] is 1L && untyped["b"] is List<object?> { Count: 3 }, "Untyped values were not resolved.");

// A typed section mixed into a loaded document.
YamlDocument document = YamlDocument.Parse("# kept\nservice:\n  name: old\n  id: a\n");
document.Set("service", YamlSerializer.SerializeToNode(new ServiceConfig { Name = "new", Id = "b" }, context.ServiceConfig));
Require(document.ToString().StartsWith("# kept\nservice:\n  name: new\n", StringComparison.Ordinal), "The serialized node was not inserted.");

// Reflection is off in a NativeAOT application, and options without generated metadata say where to get it.
Require(AppContext.TryGetSwitch("OfficeIMO.Yaml.YamlSerializer.IsReflectionEnabledByDefault", out bool reflection) && !reflection,
    "Reflection-based serialization should be turned off for PublishAot.");
try {
    new YamlSerializerOptions().GetTypeInfo(typeof(ServiceConfig));
    throw new InvalidOperationException("Options without a resolver created metadata.");
} catch (NotSupportedException exception) when (exception.Message.Contains("YamlSerializerContext", StringComparison.Ordinal)) {
}

Console.WriteLine("OfficeIMO.Yaml NativeAOT smoke passed.");

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}

public enum Level { Information, Warning, Error }

[Flags]
public enum Access { None = 0, Read = 1, Write = 2 }

public sealed record Owner(string Name, string? Email = null);

public sealed class Port {
    public int Number { get; set; }
    public string? Name { get; set; }
}

[YamlDerivedType(typeof(HttpProbe), "http")]
[YamlDerivedType(typeof(TcpProbe), "tcp")]
public abstract class Probe {
}

public sealed class HttpProbe : Probe {
    public string Path { get; init; } = "/";
}

public sealed class TcpProbe : Probe {
    public int Port { get; set; }
}

public sealed class ServiceConfig {
    public string Name { get; set; } = string.Empty;
    public int? Replicas { get; set; }
    public Level Level { get; set; }
    public Access Access { get; set; }
    public TimeSpan? Timeout { get; set; }
    public Owner? Owner { get; set; }
    public Dictionary<string, string> Labels { get; set; } = new();
    public List<Port> Ports { get; set; } = new();
    public List<Probe> Probes { get; set; } = new();
    public DateTimeOffset? Created { get; set; }
    public required string Id { get; init; }
}

[YamlSerializable(typeof(ServiceConfig))]
internal sealed partial class SmokeYamlContext : YamlSerializerContext {
}
