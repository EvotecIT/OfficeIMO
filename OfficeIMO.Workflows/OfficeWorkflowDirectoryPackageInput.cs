namespace OfficeIMO.Workflows;

/// <summary>A provider directory treated as one document by a registered package converter.</summary>
public sealed class OfficeWorkflowDirectoryPackageInput {
    /// <summary>Creates bounded package access without reading the provider during request validation.</summary>
    /// <param name="name">Selected document name, including the format extension.</param>
    /// <param name="directory">Reopenable member enumeration and permission-scoped file access.</param>
    /// <param name="sourcePublicationGuard">Verifies the selected root identity and rejects destinations inside, containing, or aliasing the source package. It must acquire any permissions needed for these checks.</param>
    /// <param name="maximumEntries">Maximum files and directories admitted before the format owner applies its own limits.</param>
    public OfficeWorkflowDirectoryPackageInput(string name, OfficeWorkflowDirectoryInput directory,
        IOfficeWorkflowPublicationGuard sourcePublicationGuard, int maximumEntries = 10_000) {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 4096 || name.IndexOfAny(['/', '\\', '\0']) >= 0)
            throw new ArgumentException("A provider package filename is required.", nameof(name));
        if (maximumEntries < 1) throw new ArgumentOutOfRangeException(nameof(maximumEntries));
        Name = name;
        Directory = directory ?? throw new ArgumentNullException(nameof(directory));
        SourcePublicationGuard = sourcePublicationGuard ?? throw new ArgumentNullException(nameof(sourcePublicationGuard));
        MaximumEntries = maximumEntries;
    }

    /// <summary>Selected filename used for routing and source diagnostics.</summary>
    public string Name { get; }
    /// <summary>Provider enumeration, re-read before publication to verify membership and content.</summary>
    public OfficeWorkflowDirectoryInput Directory { get; }
    /// <summary>Permission-aware root identity and output-separation verification.</summary>
    public IOfficeWorkflowPublicationGuard SourcePublicationGuard { get; }
    /// <summary>Maximum number of provider members, including directories.</summary>
    public int MaximumEntries { get; }
}
