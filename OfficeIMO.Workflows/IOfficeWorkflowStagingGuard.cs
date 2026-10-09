namespace OfficeIMO.Workflows;

/// <summary>Optional host policy checked before creating or writing a local destination staging directory.</summary>
/// <remarks>Implement alongside IOfficeWorkflowPublicationGuard when the host limits writable roots.
/// Checks run before and after directory creation; they do not lock filesystem identities against concurrent changes.</remarks>
public interface IOfficeWorkflowStagingGuard {
    /// <summary>Throws when staging in the absolute directory is unauthorized. Must honor cancellation.</summary>
    ValueTask EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken);
}
