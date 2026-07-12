using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>Opens the associated document in the operating system's registered application.</summary>
        public void OpenInApplication(string? filePath = null) {
            string target = string.IsNullOrEmpty(filePath) ? _filePath ?? string.Empty : filePath!;
            if (string.IsNullOrEmpty(target)) {
                throw new InvalidOperationException("The document has no associated file path.");
            }
            OfficeIMO.Core.OfficeFileLauncher.Open(target);
        }

        /// <summary>Saves the document to the path specified when created.</summary>
        public void Save() {
            ThrowIfInvalidForSave();
            if (string.IsNullOrEmpty(_filePath)) {
                if (_sourceStream == null) {
                    throw new InvalidOperationException("File path is not set.");
                }
                Save(_sourceStream);
                return;
            }
            SaveInternal(_filePath!);
        }

        /// <summary>Saves the document to a specified file path.</summary>
        public void Save(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            ThrowIfInvalidForSave();
            SaveInternal(filePath);
            _filePath = filePath;
        }

        /// <summary>Saves the document to a specified stream.</summary>
        public void Save(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            ThrowIfInvalidForSave();
            SaveInternal(stream);
        }

        /// <summary>Encodes the document as a VSDX package.</summary>
        public byte[] ToBytes() {
            ThrowIfInvalidForSave();
            using var stream = new MemoryStream();
            SaveInternal(stream);
            return stream.ToArray();
        }

        /// <summary>Encodes the document in a new writable memory stream positioned at the beginning.</summary>
        public MemoryStream ToStream() => new MemoryStream(ToBytes());

        /// <summary>Asynchronously saves to the associated path or stream.</summary>
        public Task SaveAsync(CancellationToken cancellationToken = default) {
            if (!string.IsNullOrEmpty(_filePath)) return SaveAsync(_filePath!, cancellationToken);
            if (_sourceStream != null) return SaveAsync(_sourceStream, cancellationToken);
            throw new InvalidOperationException("The document has no associated destination. Use SaveAsync(string) or SaveAsync(Stream).");
        }

        /// <summary>Asynchronously saves to a file and associates it with subsequent saves.</summary>
        public async Task SaveAsync(string filePath, CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = ToBytes();
            await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            _filePath = filePath;
        }

        /// <summary>Asynchronously saves to a caller-owned writable stream.</summary>
        public async Task SaveAsync(Stream stream, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            cancellationToken.ThrowIfCancellationRequested();
            await OfficeStreamWriter.WriteAllBytesAsync(stream, ToBytes(), cancellationToken)
                .ConfigureAwait(false);
        }

        // Wrapper to call the core implementation (kept in another partial).
        private void SaveInternal(string filePath) => SaveInternalCore(filePath);
        private void SaveInternal(Stream stream) => SaveInternalCore(stream);
    }
}
