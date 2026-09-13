using System.Collections.ObjectModel;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>Read-only native-container facts. Presence does not imply semantic or execution support.</summary>
public sealed class ProjectNativeInfo {
    internal ProjectNativeInfo(string? producer, OfficeCompoundFile compound) {
        ProducerVersion = producer;
        Profile = ProjectNativeProfile.Detect(compound);
        IsTemplate = compound.Streams.TryGetValue("\u0001CompObj", out var identifier)
            && OfficeOleCompoundObjectReader.ClipboardFormat(identifier) == "MSProject.MPT" + Profile.Version;
        Streams = new ReadOnlyCollection<ProjectNativeStreamInfo>(compound.Streams.OrderBy(s => s.Key, StringComparer.Ordinal)
            .Select(s => new ProjectNativeStreamInfo(s.Key, s.Value.Length)).ToArray());
        HasMacroStorage = compound.Entries.Any(e => e.Name.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0 || e.Name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase));
        HasSignatureStorage = Streams.Any(s => s.Path.IndexOf("signature", StringComparison.OrdinalIgnoreCase) >= 0);
        HasEmbeddedContent = compound.Entries.Any(e => e.Name.Equals("ObjectPool", StringComparison.OrdinalIgnoreCase) || e.Name.Equals("Attachments", StringComparison.OrdinalIgnoreCase));
    }
    /// <summary>The producer version recorded in Props14, when available.</summary>
    public string? ProducerVersion { get; }
    /// <summary>Detected native generation.</summary>
    public string Generation => Profile.Generation;
    internal ProjectNativeProfile Profile { get; }
    /// <summary>True when the native clipboard-format identifier declares an MPT14 document template.</summary>
    public bool IsTemplate { get; }
    /// <summary>Original stream paths and lengths, without exposing mutable bytes.</summary>
    public IReadOnlyList<ProjectNativeStreamInfo> Streams { get; }
    /// <summary>Whether conventional VBA storage names were found; macros are never executed.</summary>
    public bool HasMacroStorage { get; }
    /// <summary>Whether conventional signature stream names were found; signatures are not cryptographically verified.</summary>
    public bool HasSignatureStorage { get; }
    /// <summary>Whether conventional embedded-content storage names were found; content is not activated.</summary>
    public bool HasEmbeddedContent { get; }
}

/// <summary>A native source stream's identity and bounded size.</summary>
public sealed class ProjectNativeStreamInfo {
    internal ProjectNativeStreamInfo(string path, int length) { Path = path; Length = length; }
    /// <summary>Compound storage path.</summary>
    public string Path { get; }
    /// <summary>Original byte count.</summary>
    public int Length { get; }
}

internal sealed class ProjectNativeSource {
    internal const string ProducerMarker = "OfficeIMO.Project";
    internal readonly byte[] Bytes;
    internal readonly ProjectNativeInfo Info;
    internal Dictionary<string, object?> Snapshot = new Dictionary<string, object?>();
    internal long ModelRevision;
    internal bool CreatedByOfficeIMO;
    internal Guid? GeneratedIdentitySeed;
    internal IReadOnlyList<ProjectDiagnostic> UnrepresentedValues = Array.Empty<ProjectDiagnostic>();
    internal ProjectNativeSource(byte[] bytes, ProjectNativeInfo info) { Bytes = bytes; Info = info; }
}
