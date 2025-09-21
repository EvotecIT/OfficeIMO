using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>
        /// Loads an existing .vsdx file into a VisioDocument.
        /// </summary>
        public static VisioDocument Load(string filePath) => LoadCore(filePath);
    }
}

