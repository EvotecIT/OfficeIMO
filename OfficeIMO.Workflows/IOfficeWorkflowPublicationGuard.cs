namespace OfficeIMO.Workflows;

/// <summary>Checks application ownership of a final workflow publication candidate.</summary>
/// <remarks>
/// Called after artifact validation for each candidate, including numbered alternatives.
/// Implementations must support calls from worker threads and honor cancellation.
/// A false result rejects Fail/Replace publication and skips that candidate for Rename.
/// Exceptions fail the operation. This point-in-time check does not lock filesystem identities.
/// </remarks>
public interface IOfficeWorkflowPublicationGuard {
    /// <summary>Returns whether a file or complete directory may be published at the given absolute path.</summary>
    ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken);
}
