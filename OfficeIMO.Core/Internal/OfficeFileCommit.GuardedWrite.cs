using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        /// <summary>Stages an asynchronous write and replaces an existing destination only if its content still matches the caller's snapshot.</summary>
        /// <remarks>The existing atomic replacement and rollback owner validates the displaced file as well as the installed output.</remarks>
        internal static async Task WriteIfUnchangedAsync(string targetPath,
            Func<Stream, CancellationToken, Task> writer, Func<string, bool> destinationMatchesExpected,
            CancellationToken cancellationToken = default) {
            if (writer == null) throw new ArgumentNullException(nameof(writer));
            if (destinationMatchesExpected == null) throw new ArgumentNullException(nameof(destinationMatchesExpected));
            cancellationToken.ThrowIfCancellationRequested();
            string destination = GetFullTargetPath(targetPath);
            if (!destinationMatchesExpected(destination)) throw DestinationChanged();
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(destination, FileOptions.Asynchronous, out temporaryPath)) {
                    await writer(stream, cancellationToken).ConfigureAwait(false);
                    await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                }
                cancellationToken.ThrowIfCancellationRequested();
                string outputIdentity = ComputeFileIdentity(temporaryPath);
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(temporaryPath, destination,
                    destinationMatchesExpected,
                    installed => string.Equals(outputIdentity, ComputeFileIdentity(installed), StringComparison.Ordinal))) {
                    throw DestinationChanged();
                }
                temporaryPath = string.Empty;
            } finally {
                DeleteIfExists(temporaryPath);
            }
        }

        private static IOException DestinationChanged() => new IOException(
            "The destination changed since it was opened. Save to a different path to preserve both versions.");
    }
}
