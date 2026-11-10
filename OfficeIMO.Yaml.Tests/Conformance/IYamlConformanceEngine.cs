namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Operations a YAML engine exposes to the conformance harness.</summary>
[Flags]
public enum YamlEngineFeatures {
    /// <summary>No operation is available.</summary>
    None = 0,
    /// <summary><see cref="IYamlConformanceEngine.Parse"/> reports the event stream.</summary>
    Events = 1,
    /// <summary><see cref="IYamlConformanceEngine.Load"/> builds documents with the YAML 1.2 core schema.</summary>
    Load = 2,
    /// <summary><see cref="IYamlConformanceEngine.RoundTrip"/> loads and saves a stream without edits.</summary>
    RoundTrip = 4,
    /// <summary>The engine claims that an unedited round-trip returns the input bytes unchanged.</summary>
    Lossless = 8,
    /// <summary><see cref="IYamlConformanceEngine.Emit"/> writes an event stream as YAML text.</summary>
    Emit = 16
}

/// <summary>
/// Adapter between one YAML engine and the conformance harness.
/// Inputs are raw bytes: detecting UTF-8, UTF-16 or UTF-32 and handling a byte order mark is the engine's job.
/// Every operation rejects invalid YAML by throwing; any exception counts as a rejection.
/// </summary>
public interface IYamlConformanceEngine {
    /// <summary>Short name used in reports.</summary>
    string Name { get; }

    /// <summary>Operations this adapter implements. The harness skips checks for absent features.</summary>
    YamlEngineFeatures Features { get; }

    /// <summary>Parses a complete stream into events, from stream start to stream end.</summary>
    IReadOnlyList<YamlTestEvent> Parse(byte[] yaml);

    /// <summary>
    /// Loads every document of the stream, resolving aliases and untagged plain scalars with the YAML 1.2 core schema.
    /// An empty document loads as a <see cref="YamlConformanceNodeKind.Null"/> node.
    /// </summary>
    IReadOnlyList<YamlConformanceNode> Load(byte[] yaml);

    /// <summary>Loads the stream and saves it again without changes, returning the saved bytes.</summary>
    byte[] RoundTrip(byte[] yaml);

    /// <summary>Writes an event stream, as <see cref="Parse"/> reports it, as a new YAML document and returns its bytes.</summary>
    byte[] Emit(IReadOnlyList<YamlTestEvent> events);
}
