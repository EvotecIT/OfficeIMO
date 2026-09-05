using System;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        /// <summary>Controls Unix access permissions on an atomically published file.</summary>
        public enum UnixFileAccessPolicy {
            /// <summary>Preserves an existing destination's mode, or uses normal creation permissions.</summary>
            PreserveDestinationOrDefault,
            /// <summary>Applies owner read/write permissions to staging before publication.</summary>
            OwnerOnly
        }

        /// <summary>
        /// Atomically publishes a completed staging file with an explicit Unix access policy.
        /// Owner-only mode is applied before publication for both new and existing destinations.
        /// Windows continues to use the normal filesystem ACL behavior.
        /// </summary>
        public static void CommitTemporaryFileAtomically(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy,
            UnixFileAccessPolicy unixFileAccessPolicy) {
            if (unixFileAccessPolicy != UnixFileAccessPolicy.PreserveDestinationOrDefault &&
                unixFileAccessPolicy != UnixFileAccessPolicy.OwnerOnly) {
                throw new ArgumentOutOfRangeException(nameof(unixFileAccessPolicy));
            }
            CommitTemporaryFileCore(temporaryPath, targetPath, conflictPolicy,
                allowNonAtomicReplacementFallback: false,
                allowReadOnlyUnixDestination: false,
                ownerOnlyUnixPermissions: unixFileAccessPolicy == UnixFileAccessPolicy.OwnerOnly);
        }
    }
}
