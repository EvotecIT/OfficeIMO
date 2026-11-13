using OfficeIMO;
using OfficeIMO.Toml;

// Serialization through generated metadata only: no reflection, so trimming and NativeAOT keep every member it needs.
var context = new SmokeTomlContext(new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower });

var manifest = new Manifest {
    Package = new Package { Name = "demo", Version = "1.2.0", Edition = Edition.E2021, Released = new DateOnly(2026, 11, 10) },
    Dependencies = {
        ["serde"] = new Dependency { Version = "1", Features = new List<string> { "derive" } },
        ["local"] = new Dependency { Path = "../local" }
    },
    Bin = { new Target { Name = "cli", Path = "src/main.rs" } },
    Checks = { new HttpCheck { Path = "/health" }, new TcpCheck { Port = 5432 } },
    Owner = new Owner("Ada", "ada@example.com"),
    Built = new DateTimeOffset(2026, 11, 10, 9, 0, 0, TimeSpan.Zero),
    Ratio = double.PositiveInfinity
};

const string expected = "built = 2026-11-10T09:00:00Z\nratio = inf\n"
    + "\n[package]\nname = \"demo\"\nversion = \"1.2.0\"\nedition = \"E2021\"\nreleased = 2026-11-10\n"
    + "\n[dependencies]\nserde = { version = \"1\", features = [\"derive\"] }\nlocal = { path = \"../local\" }\n"
    + "\n[[bin]]\nname = \"cli\"\npath = \"src/main.rs\"\n"
    + "\n[[checks]]\n\"$type\" = \"http\"\npath = \"/health\"\n"
    + "\n[[checks]]\n\"$type\" = \"tcp\"\nport = 5432\n"
    + "\n[owner]\nname = \"Ada\"\nemail = \"ada@example.com\"\n";
string toml = TomlSerializer.Serialize(manifest, context.Manifest);
Require(toml == expected, "The generated serializer wrote unexpected TOML:\n" + toml);

Manifest copy = TomlSerializer.Deserialize(toml, context.Manifest);
Require(copy.Owner == manifest.Owner && copy.Package.Released == manifest.Package.Released && copy.Dependencies["local"].Path == "../local"
    && copy.Checks[0] is HttpCheck { Path: "/health" } && copy.Checks[1] is TcpCheck { Port: 5432 } && double.IsPositiveInfinity(copy.Ratio),
    "The generated serializer lost a value.");
Require(TomlSerializer.Serialize(copy, context.Manifest) == toml, "The generated serializer did not round-trip.");

// A typed section of a loaded document, and untyped values.
using TomlDocument document = TomlDocument.Parse("[tool.demo]\nname = \"x\"\nversion = \"2\"\n");
Package section = TomlSerializer.Deserialize(document.Select("tool.demo")!, context.Package);
Require(section.Name == "x" && section.Version == "2", "The document section was not read.");
var untyped = (Dictionary<string, object?>)TomlSerializer.Deserialize("a = 1\nb = [true, 2.5, \"text\"]\n", typeof(object), context);
Require(untyped["a"] is 1L && untyped["b"] is List<object?> { Count: 3 }, "Untyped values were not resolved.");

// Errors carry their position.
try {
    TomlSerializer.Deserialize("[package]\nname = 1\n", context.Manifest);
    throw new InvalidOperationException("A type mismatch was not reported.");
} catch (TomlException exception) when (exception.Code == TomlErrorCodes.TypeMismatch && exception.Line == 2) {
}

// Reflection is off in a NativeAOT application, and options without generated metadata say where to get it.
Require(AppContext.TryGetSwitch("OfficeIMO.Toml.TomlSerializer.IsReflectionEnabledByDefault", out bool reflection) && !reflection,
    "Reflection-based serialization should be turned off for PublishAot.");
try {
    new TomlSerializerOptions().GetTypeInfo(typeof(Manifest));
    throw new InvalidOperationException("Options without a resolver created metadata.");
} catch (NotSupportedException exception) when (exception.Message.Contains("TomlSerializerContext", StringComparison.Ordinal)) {
}

Console.WriteLine("OfficeIMO.Toml NativeAOT smoke passed.");

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}

public enum Edition { E2018, E2021, E2024 }

public sealed record Owner(string Name, string? Email = null);

public sealed class Package {
    public string Name { get; set; } = string.Empty;
    public required string Version { get; init; }
    public Edition Edition { get; set; }
    public DateOnly? Released { get; set; }
}

[TomlInlineTable]
public sealed class Dependency {
    public string? Version { get; set; }
    public string? Path { get; set; }
    public List<string>? Features { get; set; }
}

public sealed class Target {
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
}

[OfficeDerivedType(typeof(HttpCheck), "http")]
[OfficeDerivedType(typeof(TcpCheck), "tcp")]
public abstract class Check {
}

public sealed class HttpCheck : Check {
    public string Path { get; init; } = "/";
}

public sealed class TcpCheck : Check {
    public int Port { get; set; }
}

public sealed class Manifest {
    public Package Package { get; set; } = new Package { Version = "0.1.0" };
    public Dictionary<string, Dependency> Dependencies { get; set; } = new();
    public List<Target> Bin { get; set; } = new();
    public List<Check> Checks { get; set; } = new();
    public Owner? Owner { get; set; }
    public DateTimeOffset Built { get; set; }
    public double Ratio { get; set; }
}

[TomlSerializable(typeof(Manifest))]
[TomlSerializable(typeof(Package))]
internal sealed partial class SmokeTomlContext : TomlSerializerContext {
}
