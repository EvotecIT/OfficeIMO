using System;
using System.IO;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        /// <summary>
        /// Deletes a matching snapshot after atomically moving the current directory entry to
        /// a private same-directory quarantine. A newer entry published at the original path
        /// is never deleted. A mismatching snapshot is restored without replacing another save.
        /// </summary>
        /// <remarks>
        /// The validator receives the displaced candidate's private path, not the original
        /// destination. Missing destinations succeed. A mismatch returns false after restoration;
        /// if restoration is occupied or fails, an IOException identifies the retained candidate.
        /// </remarks>
        public static bool TryDeleteIfUnchanged(string targetPath, Func<string, bool> candidateMatchesExpected) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(candidateMatchesExpected);
#else
            if (candidateMatchesExpected == null) throw new ArgumentNullException(nameof(candidateMatchesExpected));
#endif
            string destination = GetFullTargetPath(targetPath);
            string quarantine = CreateTemporaryPath(destination);
            try { File.Move(destination, quarantine); }
            catch (FileNotFoundException) { return true; }
            catch (DirectoryNotFoundException) { return true; }

            bool restoring = false;
            try {
                if (!candidateMatchesExpected(quarantine)) {
                    restoring = true;
                    RestoreQuarantinedDestination(destination, quarantine);
                    return false;
                }
                File.Delete(quarantine);
                return true;
            } catch (Exception deletionError) {
                if (!restoring) {
                    try { RestoreQuarantinedDestination(destination, quarantine); }
                    catch (Exception restoreError) {
                        throw new IOException(
                            "The guarded deletion failed and its displaced file could not be restored. " +
                            "It remains recoverable at '" + quarantine + "'.",
                            new AggregateException(deletionError, restoreError));
                    }
                }
                throw;
            }
        }

        private static void RestoreQuarantinedDestination(string destination, string quarantine) {
            bool containsTemporary = true, preserveQuarantine = false;
            try {
                // The existing rollback owner uses a non-overwriting move. An intervening
                // save keeps its original name and the displaced candidate stays recoverable.
                RestoreMissingDisplacedDestination(destination, quarantine, ref containsTemporary, ref preserveQuarantine);
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                throw new IOException(
                    "The destination could not be restored without replacing another file. " +
                    "The displaced file remains recoverable at '" + quarantine + "'.", error);
            }
        }
    }
}
