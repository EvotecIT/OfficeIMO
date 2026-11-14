using System.Numerics;
using OfficeIMO.Toml;

namespace OfficeIMO.Toml.Tests.Serialization;

// The representative model set: each shape the serializer supports appears at least once, and the same types are read
// and written through generated and reflection metadata.

public enum Edition { E2015, E2018, E2021, E2024 }

[Flags]
public enum Features { None = 0, Std = 1, Alloc = 2, Serde = 4, Default = Std | Alloc }

/// <summary>A Cargo manifest: sections, a dictionary of inline tables, arrays of tables and kebab-case keys.</summary>
public sealed class CargoManifest {
    public CargoPackage Package { get; set; } = new CargoPackage();
    public Dictionary<string, Dependency> Dependencies { get; set; } = new Dictionary<string, Dependency>();
    [OfficePropertyName("dev-dependencies")]
    public Dictionary<string, Dependency>? DevDependencies { get; set; }
    public Dictionary<string, List<string>> Features { get; set; } = new Dictionary<string, List<string>>();
    public List<BinTarget> Bin { get; set; } = new List<BinTarget>();
    public CargoProfile? Profile { get; set; }
}

public sealed class CargoPackage {
    public string Name { get; set; } = string.Empty;
    public string Version { get; set; } = "0.1.0";
    public Edition Edition { get; set; }
    public List<string> Authors { get; set; } = new List<string>();
    public string? Description { get; set; }
    public int? RustVersion { get; set; }
    public bool Publish { get; set; } = true;
}

[TomlInlineTable]
public sealed class Dependency {
    public string? Version { get; set; }
    public string? Path { get; set; }
    public List<string>? Features { get; set; }
    public bool Optional { get; set; }
}

public sealed class BinTarget {
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public Dictionary<string, string>? Env { get; set; }
}

public sealed class CargoProfile {
    public ProfileSettings? Release { get; set; }
    public ProfileSettings? Dev { get; set; }
}

public sealed class ProfileSettings {
    public int OptLevel { get; set; }
    public bool Lto { get; set; }
    public double? CodegenUnits { get; set; }
}

/// <summary>Every scalar the serializer maps, including the four TOML date-time kinds.</summary>
public sealed class Scalars {
    public string Text { get; set; } = string.Empty;
    public char Letter { get; set; }
    public bool Flag { get; set; }
    public byte Byte { get; set; }
    public sbyte SByte { get; set; }
    public short Short { get; set; }
    public ushort UShort { get; set; }
    public int Int { get; set; }
    public uint UInt { get; set; }
    public long Long { get; set; }
    public ulong ULong { get; set; }
    public BigInteger Big { get; set; }
    public float Single { get; set; }
    public double Double { get; set; }
    public decimal Decimal { get; set; }
    public DateTimeOffset Instant { get; set; }
    public DateTime Local { get; set; }
#if NET8_0_OR_GREATER
    public DateOnly Day { get; set; }
    public TimeOnly Clock { get; set; }
#endif
    public TimeSpan Elapsed { get; set; }
    public TomlDateTime Raw { get; set; }
    public Guid Id { get; set; }
    public Uri? Link { get; set; }
    public byte[]? Blob { get; set; }
    public Features Flags { get; set; }
    public Edition? MaybeEdition { get; set; }
}

public sealed class Collections {
    public int[] Numbers { get; set; } = Array.Empty<int>();
    public List<List<int>> Matrix { get; set; } = new List<List<int>>();
    public HashSet<string> Unique { get; set; } = new HashSet<string>();
    public IReadOnlyList<double> Readings { get; set; } = Array.Empty<double>();
    public Dictionary<int, string> ById { get; set; } = new Dictionary<int, string>();
    public Dictionary<Edition, bool> ByEdition { get; set; } = new Dictionary<Edition, bool>();
    public List<Point> Points { get; set; } = new List<Point>();
    [TomlInlineTable]
    public Point Origin { get; set; }
    public object? Anything { get; set; }
}

public struct Point {
    public int X { get; set; }
    public int Y { get; set; }
}

public sealed record Person(string Name, int Age, string? Email = null);

public sealed class Team {
    public string Name { get; set; } = string.Empty;
    public List<Person> Members { get; set; } = new List<Person>();
}

public sealed class Account {
    public required string Id { get; init; }
    public string Owner { get; init; } = "unknown";
    [OfficeRequired]
    public decimal Balance { get; set; }
}

public readonly struct Size {
    [OfficeConstructor]
    public Size(int width, int height) {
        Width = width;
        Height = height;
    }

    public int Width { get; }
    public int Height { get; }
}

[OfficeDerivedType(typeof(HttpProbe), "http")]
[OfficeDerivedType(typeof(TcpProbe), "tcp")]
public abstract class Probe {
    public int Interval { get; set; }
}

public sealed class HttpProbe : Probe {
    public string Path { get; set; } = "/";
}

public sealed class TcpProbe : Probe {
    public int Port { get; set; }
}

[OfficePolymorphic(DiscriminatorPropertyName = "kind")]
[OfficeDerivedType(typeof(Circle), "circle")]
public class Shape {
    public string Color { get; set; } = "black";
}

public sealed class Circle : Shape {
    public double Radius { get; set; }
}

public sealed class Monitoring {
    public Probe? Probe { get; set; }
    public List<Probe> Probes { get; set; } = new List<Probe>();
    public List<Shape> Shapes { get; set; } = new List<Shape>();
    public Size Window { get; set; }
}

public sealed class Node {
    public string Name { get; set; } = string.Empty;
    public Node? Child { get; set; }
}

public sealed class RuffSettings {
    public int LineLength { get; set; } = 88;
    public string? TargetVersion { get; set; }
    public RuffLint Lint { get; set; } = new RuffLint();
}

public sealed class RuffLint {
    public List<string> Select { get; set; } = new List<string>();
    public List<string> Ignore { get; set; } = new List<string>();
}

[TomlSerializable(typeof(CargoManifest))]
[TomlSerializable(typeof(Scalars))]
[TomlSerializable(typeof(Collections))]
[TomlSerializable(typeof(Team))]
[TomlSerializable(typeof(Account))]
[TomlSerializable(typeof(Monitoring))]
[TomlSerializable(typeof(Node))]
[TomlSerializable(typeof(RuffSettings))]
[TomlSerializable(typeof(Dictionary<string, object>))]
public partial class TestTomlContext : TomlSerializerContext {
}
