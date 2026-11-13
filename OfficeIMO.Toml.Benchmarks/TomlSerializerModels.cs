namespace OfficeIMO.Toml.Benchmarks;

// Models of the serializer workloads. Both libraries map them with a kebab-case naming policy and skip the keys a model
// does not declare, so each reads exactly the same members.

/// <summary>A PEP 621 <c>pyproject.toml</c>.</summary>
public sealed class PyProject {
    public BuildSystem? BuildSystem { get; set; }
    public PythonProject? Project { get; set; }
    public Dictionary<string, object>? Tool { get; set; }
}

public sealed class BuildSystem {
    public List<string> Requires { get; set; } = new();
    public string? BuildBackend { get; set; }
}

public sealed class PythonProject {
    public string Name { get; set; } = string.Empty;
    public string? Version { get; set; }
    public string? Description { get; set; }
    public string? Readme { get; set; }
    public string? RequiresPython { get; set; }
    public Dictionary<string, string>? License { get; set; }
    public List<string>? Keywords { get; set; }
    public List<PythonPerson>? Authors { get; set; }
    public List<PythonPerson>? Maintainers { get; set; }
    public List<string>? Classifiers { get; set; }
    public List<string>? Dependencies { get; set; }
    public Dictionary<string, List<string>>? OptionalDependencies { get; set; }
    public Dictionary<string, string>? Urls { get; set; }
    public Dictionary<string, string>? Scripts { get; set; }
}

public sealed class PythonPerson {
    public string? Name { get; set; }
    public string? Email { get; set; }
}

/// <summary>A Cargo workspace manifest: hundreds of inline-table dependencies, lint tables and build profiles.</summary>
public sealed class CargoWorkspaceManifest {
    public CargoWorkspace? Workspace { get; set; }
    public Dictionary<string, CargoBuildProfile>? Profile { get; set; }
}

public sealed class CargoWorkspace {
    public List<string> Members { get; set; } = new();
    public List<string>? Exclude { get; set; }
    public string? Resolver { get; set; }
    public CargoWorkspacePackage? Package { get; set; }
    public Dictionary<string, CargoDependency> Dependencies { get; set; } = new();
    public Dictionary<string, Dictionary<string, object>>? Lints { get; set; }
}

public sealed class CargoWorkspacePackage {
    public string? Edition { get; set; }
    public string? RustVersion { get; set; }
    public string? Homepage { get; set; }
    public string? Repository { get; set; }
    public List<string>? Authors { get; set; }
    public string? License { get; set; }
}

public sealed class CargoDependency {
    public string? Version { get; set; }
    public string? Path { get; set; }
    public string? Git { get; set; }
    public string? Rev { get; set; }
    public string? Branch { get; set; }
    public string? Tag { get; set; }
    public string? Package { get; set; }
    public bool? DefaultFeatures { get; set; }
    public bool? Optional { get; set; }
    public List<string>? Features { get; set; }
}

public sealed class CargoBuildProfile {
    public string? Inherits { get; set; }
    public object? Debug { get; set; }
    public object? Strip { get; set; }
    public object? Lto { get; set; }
    public object? OptLevel { get; set; }
    public long? CodegenUnits { get; set; }
    public string? Panic { get; set; }
}

/// <summary>A <c>uv.lock</c>: thousands of package tables with inline-table sources, artifacts and dependencies.</summary>
public sealed class UvLock {
    public long Version { get; set; }
    public long Revision { get; set; }
    public string? RequiresPython { get; set; }
    public List<string>? ResolutionMarkers { get; set; }
    public List<LockPackage> Package { get; set; } = new();
}

public sealed class LockPackage {
    public string Name { get; set; } = string.Empty;
    public string? Version { get; set; }
    public Dictionary<string, string>? Source { get; set; }
    public List<LockDependency>? Dependencies { get; set; }
    public Dictionary<string, List<LockDependency>>? OptionalDependencies { get; set; }
    public List<string>? ResolutionMarkers { get; set; }
    public LockArtifact? Sdist { get; set; }
    public List<LockArtifact>? Wheels { get; set; }
}

public sealed class LockDependency {
    public string Name { get; set; } = string.Empty;
    public string? Marker { get; set; }
    public string? Specifier { get; set; }
    public List<string>? Extra { get; set; }
}

public sealed class LockArtifact {
    public string? Url { get; set; }
    public string? Path { get; set; }
    public string? Hash { get; set; }
    public long? Size { get; set; }
    public string? UploadTime { get; set; }
}

[TomlSerializable(typeof(PyProject))]
[TomlSerializable(typeof(CargoWorkspaceManifest))]
[TomlSerializable(typeof(UvLock))]
public sealed partial class OfficeIMOBenchmarkContext : TomlSerializerContext {
}
