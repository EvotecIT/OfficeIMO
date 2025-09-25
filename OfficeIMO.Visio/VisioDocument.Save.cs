using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>Saves the document to the path specified when created.</summary>
        public void Save() {
            if (string.IsNullOrEmpty(_filePath)) {
                throw new InvalidOperationException("File path is not set.");
            }
            SaveInternal(_filePath!);
        }

        /// <summary>Saves the document to a specified file path.</summary>
        public void Save(string filePath) {
            _filePath = filePath;
            SaveInternal(filePath);
        }

        // Wrapper to call the core implementation (kept in another partial).
        private void SaveInternal(string filePath) => SaveInternalCore(filePath);
    }
}

