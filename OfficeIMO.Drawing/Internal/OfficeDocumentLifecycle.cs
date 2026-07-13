using System;
using System.IO;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Applies the lifecycle and associated-destination rules shared by Office document packages.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeDocumentLifecycle {
        /// <summary>Rejects lifecycle combinations that cannot persist safely.</summary>
        public static void Validate(
            DocumentAccessMode accessMode,
            DocumentPersistenceMode persistenceMode,
            string subject) {
            if (accessMode == DocumentAccessMode.ReadOnly &&
                persistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException($"A read-only {subject} cannot use SaveOnDispose persistence.");
            }
        }

        /// <summary>Validates a stream that will always be used as an associated destination.</summary>
        public static void EnsureAssociatedDestination(Stream stream, string parameterName) {
            if (stream == null) throw new ArgumentNullException(parameterName);
            if (!stream.CanWrite) {
                throw new ArgumentException("Stream must be writable.", parameterName);
            }
            if (!OfficeStreamWriter.CanReplaceContents(stream)) {
                throw new ArgumentException(
                    "Stream must support seeking when used as an associated destination.",
                    parameterName);
            }
        }

        /// <summary>Validates a source stream when disposal must copy the complete artifact back to it.</summary>
        public static void EnsureSaveOnDisposeDestination(
            Stream stream,
            DocumentPersistenceMode persistenceMode,
            string parameterName) {
            if (persistenceMode != DocumentPersistenceMode.SaveOnDispose) return;
            if (!stream.CanWrite) {
                throw new ArgumentException(
                    "Stream must be writable when SaveOnDispose is enabled.",
                    parameterName);
            }
            if (!stream.CanSeek) {
                throw new ArgumentException(
                    "Stream must support seeking when SaveOnDispose is enabled.",
                    parameterName);
            }
        }

        /// <summary>
        /// Returns the caller-owned stream only when it can safely be retained as an editable associated destination.
        /// </summary>
        public static Stream? ResolveAssociatedDestination(
            Stream? stream,
            DocumentAccessMode accessMode) =>
            accessMode == DocumentAccessMode.ReadWrite &&
            stream != null &&
            OfficeStreamWriter.CanReplaceContents(stream)
                ? stream
                : null;
    }
}
