using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>Resolves named children of a selected provider folder for verified direct publication.</summary>
/// <remarks>The resolver must not create or modify items. Creation belongs in the returned output's write
/// factory or deferred preparation, which run after durable recovery and authorization. Files use Replace semantics individually;
/// previously verified writes cannot be rolled back when a later file fails.</remarks>
public sealed class OfficeWorkflowDirectoryOutput {
    /// <summary>Creates a folder adapter with a non-mutating child resolver.</summary>
    public OfficeWorkflowDirectoryOutput(Func<string, CancellationToken, Task<OfficeWorkflowDirectoryOutputFile>> resolveFile) =>
        ResolveFile = resolveFile ?? throw new ArgumentNullException(nameof(resolveFile));

    /// <summary>Gets the resolver for one filename, preserving its name and the selected parent.</summary>
    public Func<string, CancellationToken, Task<OfficeWorkflowDirectoryOutputFile>> ResolveFile { get; }
}

/// <summary>A child destination and its provider access.</summary>
public sealed class OfficeWorkflowDirectoryOutputFile {
    /// <summary>Creates a descriptor without modifying the destination.</summary>
    public OfficeWorkflowDirectoryOutputFile(string location, OfficeWorkflowStreamOutput output) {
        Location = OfficeStorageIdentity.Normalize(location);
        Output = output ?? throw new ArgumentNullException(nameof(output));
    }
    /// <summary>Gets the existing child location, or the selected parent when deferred preparation creates a new child.</summary>
    public string Location { get; }
    /// <summary>Gets the verified stream destination.</summary>
    public OfficeWorkflowStreamOutput Output { get; }
}
