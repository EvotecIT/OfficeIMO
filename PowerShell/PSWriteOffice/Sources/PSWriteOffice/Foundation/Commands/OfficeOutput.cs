using System;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace PSWriteOffice.Foundation;

/// <summary>Applies shared overwrite policy to a completed Export-/Convert- or New- file write.</summary>
/// <remarks>Call after provider resolution and ShouldProcess on the pipeline thread.</remarks>
internal static class OfficeOutput {
    internal static void Write(string path, Action<Stream> writer, bool noClobber, CancellationToken cancellationToken = default) {
        OfficePath.ValidateDestination(path, noClobber);
        OfficeFileCommit.WriteAtomically(path, writer, cancellationToken,
            noClobber ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace);
    }
    internal static void WriteNew(string path, Action<Stream> writer, bool force, CancellationToken cancellationToken = default) =>
        Write(path, writer, noClobber: !force, cancellationToken);
}
