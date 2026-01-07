using System;
using System.IO;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>Saves the document to the path specified when created.</summary>
        public void Save() {
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
            _filePath = filePath;
            SaveInternal(filePath);
        }

        /// <summary>Saves the document to a specified stream.</summary>
        public void Save(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            SaveInternal(stream);
        }

        // Wrapper to call the core implementation (kept in another partial).
        private void SaveInternal(string filePath) => SaveInternalCore(filePath);
        private void SaveInternal(Stream stream) => SaveInternalCore(stream);
    }
}

