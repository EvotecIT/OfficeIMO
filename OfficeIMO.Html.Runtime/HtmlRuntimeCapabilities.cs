namespace OfficeIMO.Html.Runtime;

/// <summary>Stable capability identifiers advertised before a runtime context starts.</summary>
public static class HtmlRuntimeCapabilityIds {
    /// <summary>A page remains alive across operations.</summary>
    public const string PersistentPage = "page.persistent";
    /// <summary>Pages run inside an isolated context.</summary>
    public const string IsolatedContext = "context.isolated";
    /// <summary>The page supports navigation.</summary>
    public const string Navigation = "page.navigation";
    /// <summary>The page supports JavaScript evaluation.</summary>
    public const string ScriptEvaluation = "script.evaluate";
    /// <summary>The page supports bounded classic-script realms in same-origin child frames.</summary>
    public const string SameOriginChildFrameRealms = "frame.same-origin-classic-realms";
    /// <summary>The page exposes bounded semantic observations.</summary>
    public const string SemanticObservation = "observation.semantic";
    /// <summary>The page exposes layout geometry observations.</summary>
    public const string VisualObservation = "observation.visual-geometry";
    /// <summary>The page exposes screenshot artifact references.</summary>
    public const string ScreenshotObservation = "observation.screenshot";
    /// <summary>Observed elements can be targeted while their revision is current.</summary>
    public const string RevisionBoundReferences = "observation.revision-bound-references";
    /// <summary>The page accepts provider-neutral structured actions.</summary>
    public const string StructuredActions = "automation.structured-actions";
    /// <summary>The page records bounded provider-neutral operation traces.</summary>
    public const string OperationTrace = "trace.operations";
    /// <summary>The trace includes bounded structured resource, console, policy, lifecycle and download events.</summary>
    public const string StructuredTraceEvents = "trace.structured-events";
    /// <summary>Captured artifacts expose deterministic content manifests.</summary>
    public const string DeterministicArtifactManifest = "artifact.deterministic-manifest";
    /// <summary>A context can own multiple concurrent pages.</summary>
    public const string MultiplePages = "context.multiple-pages";
}

/// <summary>Provider identity, profiles, limits, and supported runtime capabilities.</summary>
public sealed class HtmlRuntimeProviderDescriptor {
    /// <summary>Creates an immutable provider descriptor.</summary>
    public HtmlRuntimeProviderDescriptor(
        string id,
        string version,
        IEnumerable<HtmlRuntimeProfile> profiles,
        IEnumerable<string> capabilities,
        int maximumContexts,
        int maximumPagesPerContext) {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        ArgumentException.ThrowIfNullOrWhiteSpace(version);
        ArgumentNullException.ThrowIfNull(profiles);
        ArgumentNullException.ThrowIfNull(capabilities);
        if (maximumContexts <= 0) throw new ArgumentOutOfRangeException(nameof(maximumContexts));
        if (maximumPagesPerContext <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPagesPerContext));
        Id = id;
        Version = version;
        Profiles = Array.AsReadOnly(profiles.Distinct().OrderBy(value => value).ToArray());
        Capabilities = Array.AsReadOnly(capabilities.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray());
        MaximumContexts = maximumContexts;
        MaximumPagesPerContext = maximumPagesPerContext;
    }

    /// <summary>Stable provider identifier.</summary>
    public string Id { get; }
    /// <summary>Provider implementation version.</summary>
    public string Version { get; }
    /// <summary>Supported behavior profiles.</summary>
    public IReadOnlyList<HtmlRuntimeProfile> Profiles { get; }
    /// <summary>Stable capability identifiers.</summary>
    public IReadOnlyList<string> Capabilities { get; }
    /// <summary>Maximum contexts supported by one host.</summary>
    public int MaximumContexts { get; }
    /// <summary>Maximum live pages supported by one context.</summary>
    public int MaximumPagesPerContext { get; }
    /// <summary>Returns whether this provider advertises a capability.</summary>
    public bool Supports(string capability) => Capabilities.Contains(capability, StringComparer.Ordinal);

    internal static IReadOnlyList<HtmlRuntimeProfile> ProcessWorkerProfiles { get; } = Array.AsReadOnly(new[] {
        HtmlRuntimeProfile.ScriptedDocumentV1, HtmlRuntimeProfile.WebApplicationV1
    });
    internal static IReadOnlyList<string> ProcessWorkerCapabilities { get; } = Array.AsReadOnly(new[] {
        HtmlRuntimeCapabilityIds.PersistentPage,
        HtmlRuntimeCapabilityIds.IsolatedContext,
        HtmlRuntimeCapabilityIds.Navigation,
        HtmlRuntimeCapabilityIds.ScriptEvaluation,
        HtmlRuntimeCapabilityIds.SameOriginChildFrameRealms,
        HtmlRuntimeCapabilityIds.SemanticObservation,
        HtmlRuntimeCapabilityIds.VisualObservation,
        HtmlRuntimeCapabilityIds.RevisionBoundReferences,
        HtmlRuntimeCapabilityIds.StructuredActions,
        HtmlRuntimeCapabilityIds.OperationTrace,
        HtmlRuntimeCapabilityIds.StructuredTraceEvents,
        HtmlRuntimeCapabilityIds.DeterministicArtifactManifest
    });

    internal static HtmlRuntimeProviderDescriptor ProcessWorker { get; } = CreateProcessWorker(
        typeof(HtmlRuntimeProviderDescriptor).Assembly.GetName().Version?.ToString() ?? "unknown",
        ProcessWorkerProfiles,
        ProcessWorkerCapabilities,
        int.MaxValue,
        1);

    internal static HtmlRuntimeProviderDescriptor CreateProcessWorker(string version, IEnumerable<HtmlRuntimeProfile> profiles,
        IEnumerable<string> capabilities, int maximumContexts, int maximumPagesPerContext) => new(
            "officeimo.trusted-process", version, profiles, capabilities, maximumContexts, maximumPagesPerContext);
}
