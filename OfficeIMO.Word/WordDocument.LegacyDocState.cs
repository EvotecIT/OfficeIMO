using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private LegacyDocImportDiagnostic[] _legacyDocImportDiagnostics = Array.Empty<LegacyDocImportDiagnostic>();
        private LegacyDocUnsupportedFeature[] _legacyDocUnsupportedFeatures = Array.Empty<LegacyDocUnsupportedFeature>();
        private string? _legacyDocSourcePath;

        /// <summary>
        /// Gets whether this document was projected from a legacy binary `.doc` source.
        /// </summary>
        public bool WasLoadedFromLegacyDoc { get; private set; }

        /// <summary>
        /// Gets the legacy `.doc` source path when the document was loaded from a path.
        /// </summary>
        public string? LegacyDocSourcePath => _legacyDocSourcePath;

        /// <summary>
        /// Gets diagnostics produced while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> LegacyDocImportDiagnostics => _legacyDocImportDiagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered while importing the legacy `.doc` document.
        /// </summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> LegacyDocUnsupportedFeatures => _legacyDocUnsupportedFeatures;

        internal void MarkLoadedFromLegacyDoc(string? sourcePath, LegacyDocDocument document, bool attachSourcePathForSave = false) {
            WasLoadedFromLegacyDoc = true;
            _legacyDocSourcePath = sourcePath;
            _legacyDocImportDiagnostics = document.Diagnostics.ToArray();
            _legacyDocUnsupportedFeatures = document.UnsupportedFeatures.ToArray();
            FilePath = attachSourcePathForSave && sourcePath != null
                ? sourcePath
                : string.Empty;
        }
    }
}
