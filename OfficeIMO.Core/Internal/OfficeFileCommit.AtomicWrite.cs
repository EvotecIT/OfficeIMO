using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        /// <summary>
        /// Stages and publishes a complete file without temporarily removing an existing
        /// destination. Unsupported atomic replacement fails while preserving that destination.
        /// </summary>
        public static void WriteAtomically(string targetPath, Action<Stream> writer,
            CancellationToken cancellationToken = default, ConflictPolicy conflictPolicy = ConflictPolicy.Replace) =>
            WriteCore(targetPath, writer, cancellationToken, conflictPolicy, requireAtomicReplacement: true);

        /// <summary>
        /// Asynchronously stages and publishes a complete file without a backup-and-move
        /// fallback, observing cancellation before publication and cleaning up failed staging.
        /// </summary>
        public static Task WriteAtomicallyAsync(string targetPath, Func<Stream, CancellationToken, Task> writer,
            CancellationToken cancellationToken = default, ConflictPolicy conflictPolicy = ConflictPolicy.Replace) =>
            WriteCoreAsync(targetPath, writer, conflictPolicy, cancellationToken, requireAtomicReplacement: true);
    }
}
