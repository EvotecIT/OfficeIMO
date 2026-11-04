using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO;

/// <summary>Optional host policy checked before creating or writing a local destination staging directory.</summary>
/// <remarks>Implement alongside IOfficePublicationGuard when the host limits writable roots.
/// Checks run before and after directory creation; they do not lock filesystem identities against concurrent changes.</remarks>
public interface IOfficeStagingGuard {
    /// <summary>Throws when staging in the absolute directory is unauthorized. Must honor cancellation.</summary>
    Task EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken);
}
